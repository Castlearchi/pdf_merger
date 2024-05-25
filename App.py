import ctypes

# DPIスケーリングを設定
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception as e:
    print(f"DPIスケーリングの設定エラー: {e}")

import fitz  # PyMuPDF
import io
import os
from PIL import Image, ImageTk
import random
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

class PDFMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Merger")

        # 左フレームの作成
        self.left_frame = ttk.Frame(root, padding="10")
        self.left_frame.grid(row=0, column=0, sticky="nsew")

        # 右フレームの作成
        self.right_frame = ttk.Frame(root, padding="10")
        self.right_frame.grid(row=0, column=1, sticky="nsew")

        # メインウィンドウのグリッドを設定
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_rowconfigure(0, weight=1)

        # 左フレームにウィジェットを配置
        self.add_button = ttk.Button(
            self.left_frame, text="Add PDF", command=self.add_pdf
        )
        self.add_button.grid(row=0, column=0, sticky="ew", pady=5)

        self.file_lists_frame = tk.Frame(self.left_frame)
        self.file_lists_frame.grid(row=1, column=0, sticky="nsew", pady=5)

        scrollbar = tk.Scrollbar(self.file_lists_frame, orient=tk.HORIZONTAL)
        scrollbar.pack(side=tk.BOTTOM, fill=tk.X)

        self.file_lists_canvas = tk.Canvas(
            self.file_lists_frame, xscrollcommand=scrollbar.set
        )
        self.file_lists_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar.config(command=self.file_lists_canvas.xview)

        self.file_lists_inner_frame = tk.Frame(self.file_lists_canvas)
        self.file_lists_canvas.create_window(
            (0, 0), window=self.file_lists_inner_frame, anchor="nw"
        )

        add_to_output_button = ttk.Button(self.left_frame, text="Add to Output", command=self.add_to_output)
        add_to_output_button.grid(row=2, column=0, sticky="ew", pady=5)

        # 左フレームのグリッドを設定
        self.left_frame.grid_rowconfigure(1, weight=1)
        self.left_frame.grid_columnconfigure(0, weight=1)

        # 右フレームにウィジェットを配置
        right_label = ttk.Label(self.right_frame, text="Preview")
        right_label.grid(row=0, column=0, sticky="ew", pady=5)

        self.preview_label = tk.Label(self.right_frame, background="gray")
        self.preview_label.grid(row=1, column=0, sticky="nsew", pady=5)
        self.preview_image = None

        self.output_listbox_label = ttk.Label(self.right_frame, text="生成ファイル")
        self.output_listbox_label.grid(row=2, column=0, sticky="ew", pady=5)

        self.output_listbox = tk.Listbox(self.right_frame)
        self.output_listbox.grid(row=3, column=0, sticky="nsew", pady=5)
        self.output_listbox.bind("<Delete>", self.remove_output_page)
        vscrollbar = tk.Scrollbar(self.output_listbox, orient=tk.VERTICAL)
        vscrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        create_button = ttk.Button(self.right_frame, text="Create PDF", command=self.create_pdf)
        create_button.grid(row=4, column=0, sticky="ew", pady=5)

        # 右フレームのグリッドを設定
        self.right_frame.grid_rowconfigure(1, weight=1)
        self.right_frame.grid_rowconfigure(3, weight=1)
        self.right_frame.grid_columnconfigure(0, weight=1)

        self.pdf_files = []
        self.output_pages = []

    def add_pdf(self):
        file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if file_path:
            pdf_document = fitz.open(file_path)
            pages = []
            for page_num in range(len(pdf_document)):
                pages.append((file_path, page_num, 0))  # Add default rotation 0
            pdf_document.close()

            file_info = {
                "file_path": file_path,
                "pages": pages,
                "listbox": None,  # Placeholder for the listbox widget
            }
            self.pdf_files.append(file_info)
            self.create_file_listbox(file_info)

    def create_file_listbox(self, file_info):
        frame = tk.Frame(self.file_lists_inner_frame)
        frame.pack(side=tk.LEFT, padx=2, pady=5)
        color = self.get_random_color()
        label = tk.Label(
            frame, text=os.path.basename(file_info["file_path"]), bg=color, fg="white"
        )
        label.pack(fill=tk.X)

        listbox = tk.Listbox(frame, selectmode=tk.SINGLE, bg=color, fg="white")
        listbox.pack(fill=tk.BOTH, expand=True)
        listbox.bind("<<ListboxSelect>>", self.show_preview)
        for page_num in range(len(file_info["pages"])):
            listbox.insert(
                tk.END, f"{os.path.basename(file_info['file_path'])}-{page_num + 1}"
            )
        listbox.bind("<Delete>", self.remove_output_page)

        file_info["listbox"] = listbox  # Store the listbox in the file info

        self.file_lists_inner_frame.update_idletasks()
        self.file_lists_canvas.config(scrollregion=self.file_lists_canvas.bbox("all"))

    def get_random_color(self):
        r = lambda: random.randint(0, 255)
        return f"#{r():02x}{r():02x}{r():02x}"

    def show_preview(self, event):
        listbox = None
        selected_idx = None

        if event is not None:
            listbox = event.widget
            selected_idx = listbox.curselection()
        else:
            # Find the listbox that has a selection
            for file_info in self.pdf_files:
                if file_info["listbox"].curselection():
                    listbox = file_info["listbox"]
                    selected_idx = listbox.curselection()
                    break

        if selected_idx:
            idx = selected_idx[0]
            file_path, page_num, rotation = self.get_page_info(listbox, idx)
            pdf_document = fitz.open(file_path)
            page = pdf_document[page_num]
            pix = page.get_pixmap(
                matrix=fitz.Matrix(2, 2)
            )  # Zoom in for better quality
            pdf_document.close()

            image = Image.open(io.BytesIO(pix.tobytes()))
            if rotation:
                image = image.rotate(rotation, expand=True)

            # Resize the image to fit into the preview_label
            max_width, max_height = (
                self.preview_label.winfo_width(),
                self.preview_label.winfo_height(),
            )
            image.thumbnail((max_width, max_height))

            self.preview_image = ImageTk.PhotoImage(image)
            self.preview_label.config(image=self.preview_image)

    def get_page_info(self, listbox, idx):
        for file_info in self.pdf_files:
            if listbox == file_info["listbox"]:
                return file_info["pages"][idx]

    def add_to_output(self):
        for file_info in self.pdf_files:
            listbox = file_info["listbox"]
            selected_idx = listbox.curselection()
            if selected_idx:
                idx = selected_idx[0]
                file_path, page_num, rotation = file_info["pages"][idx]
                self.output_pages.append((file_path, page_num, rotation))
                self.output_listbox.insert(
                    tk.END, f"{os.path.basename(file_path)}-{page_num + 1}"
                )

    def remove_output_page(self, event):
        selected_idx = self.output_listbox.curselection()
        if selected_idx:
            idx = selected_idx[0]
            del self.output_pages[idx]
            self.output_listbox.delete(idx)

    def create_pdf(self):
        output_path = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")]
        )
        if output_path:
            pdf_writer = fitz.open()

            for file_path, page_num, rotation in self.output_pages:
                pdf_document = fitz.open(file_path)
                page = pdf_document.load_page(page_num)
                page.set_rotation(rotation)
                pdf_writer.insert_pdf(
                    pdf_document, from_page=page_num, to_page=page_num
                )
                pdf_document.close()

            pdf_writer.save(output_path)
            pdf_writer.close()

            messagebox.showinfo("Success", "PDF created successfully!")

if __name__ == "__main__":
    root = tk.Tk()
    root.geometry("1600x800")
    app = PDFMergerApp(root)
    root.mainloop()
