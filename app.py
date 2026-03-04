from __future__ import annotations

import shutil
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from word_merge import CATEGORY_KEYS, chunks, default_output_filename, merge_docx_files, render_template_page

APP_TITLE = "Word Photo Merge"
TEMPLATE_DIR = Path("templates")
TEMP_DIR = Path("tempfiles")


class WordMergeApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("980x680")

        TEMPLATE_DIR.mkdir(exist_ok=True)
        TEMP_DIR.mkdir(exist_ok=True)

        self.category_photos: dict[str, list[Path]] = {category: [] for category in CATEGORY_KEYS}
        self.progress_var = tk.DoubleVar(value=0)
        self.status_var = tk.StringVar(value="Hazır")

        self._cleanup_prompt_on_start()
        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self._on_exit)

    def _cleanup_prompt_on_start(self):
        existing = [p for p in TEMP_DIR.iterdir() if p.is_file()]
        if existing and messagebox.askyesno("Tempfiles", "Mevcut tempfiles bulundu. Temizlensin mi?"):
            self._clean_temp_dir()

    def _build_ui(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Header.TLabel", font=("Segoe UI", 13, "bold"))
        style.configure("Card.TLabelframe", padding=10)

        main = ttk.Frame(self.root, padding=16)
        main.pack(fill="both", expand=True)

        title = ttk.Label(main, text="Word Fotoğraf Birleştirme", style="Header.TLabel")
        title.pack(anchor="w", pady=(0, 12))

        config_frame = ttk.Labelframe(main, text="Template Ayarları", style="Card.TLabelframe")
        config_frame.pack(fill="x", pady=(0, 10))

        ttk.Label(
            config_frame,
            text=(
                "Template dosya adları: pre.docx, post.docx, teardown.docx, handle_side_cover.docx\n"
                "Her template içinde görsel yerleri {{IMG1}} ... {{IMG6}} placeholderları ile işaretlenmelidir."
            ),
        ).pack(anchor="w")

        ttk.Button(config_frame, text="Templates klasörünü aç", command=self._open_templates_info).pack(anchor="w", pady=(8, 0))

        categories_frame = ttk.Frame(main)
        categories_frame.pack(fill="both", expand=True)

        self.listboxes: dict[str, tk.Listbox] = {}
        for i, category in enumerate(CATEGORY_KEYS):
            card = ttk.Labelframe(categories_frame, text=category, style="Card.TLabelframe")
            card.grid(row=i // 2, column=i % 2, padx=8, pady=8, sticky="nsew")

            listbox = tk.Listbox(card, height=8)
            listbox.pack(fill="both", expand=True, pady=(0, 8))
            self.listboxes[category] = listbox

            button_row = ttk.Frame(card)
            button_row.pack(fill="x")
            ttk.Button(button_row, text="Fotoğraf ekle", command=lambda c=category: self._add_photos(c)).pack(side="left")
            ttk.Button(button_row, text="Temizle", command=lambda c=category: self._clear_category(c)).pack(side="left", padx=6)

        for col in range(2):
            categories_frame.columnconfigure(col, weight=1)
        for row in range(2):
            categories_frame.rowconfigure(row, weight=1)

        controls = ttk.Frame(main)
        controls.pack(fill="x", pady=(12, 6))

        ttk.Button(controls, text="Birleştir", command=self._merge).pack(side="left")
        ttk.Button(controls, text="Tempfiles temizle", command=self._clean_temp_dir).pack(side="left", padx=8)

        self.progress = ttk.Progressbar(controls, mode="determinate", variable=self.progress_var)
        self.progress.pack(side="left", fill="x", expand=True, padx=12)

        ttk.Label(controls, textvariable=self.status_var).pack(side="left")

    def _open_templates_info(self):
        messagebox.showinfo(
            "Templates",
            f"Template dosyalarını şu klasöre koyun:\n{TEMPLATE_DIR.resolve()}\n\n"
            "Dosya isimleri:\n"
            "- pre.docx\n- post.docx\n- teardown.docx\n- handle_side_cover.docx",
        )

    def _refresh_listbox(self, category: str):
        lb = self.listboxes[category]
        lb.delete(0, tk.END)
        for photo in self.category_photos[category]:
            lb.insert(tk.END, photo.name)

    def _add_photos(self, category: str):
        paths = filedialog.askopenfilenames(
            title=f"{category} için fotoğraf seç",
            filetypes=[("Image files", "*.jpg *.jpeg *.png *.bmp *.tiff")],
        )
        if not paths:
            return

        selected = [Path(p) for p in paths]
        self.category_photos[category].extend(selected)
        self._refresh_listbox(category)

        total = sum(len(v) for v in self.category_photos.values())
        self.progress.configure(maximum=max(total, 1))
        self.progress_var.set(total)
        self.status_var.set(f"{category}: {len(selected)} fotoğraf eklendi")

    def _clear_category(self, category: str):
        self.category_photos[category] = []
        self._refresh_listbox(category)

    def _clean_temp_dir(self):
        TEMP_DIR.mkdir(exist_ok=True)
        for item in TEMP_DIR.iterdir():
            if item.is_file():
                item.unlink(missing_ok=True)
            elif item.is_dir():
                shutil.rmtree(item, ignore_errors=True)
        self.status_var.set("Tempfiles temizlendi")

    def _merge(self):
        selected_categories = [c for c, photos in self.category_photos.items() if photos]
        if not selected_categories:
            messagebox.showwarning("Eksik", "En az bir kategoriye fotoğraf ekleyin")
            return

        try:
            TEMP_DIR.mkdir(exist_ok=True)
            self._clean_temp_dir()
            temp_outputs: list[Path] = []

            total_images = sum(len(self.category_photos[c]) for c in selected_categories)
            self.progress.configure(maximum=max(total_images, 1))
            current = 0

            for category in selected_categories:
                key = CATEGORY_KEYS[category]
                template_path = TEMPLATE_DIR / f"{key}.docx"
                photos = self.category_photos[category]

                for page_index, photo_chunk in enumerate(chunks(photos, 6), start=1):
                    out_doc = TEMP_DIR / f"{key}_{page_index}.docx"
                    render_template_page(template_path, photo_chunk, out_doc)
                    temp_outputs.append(out_doc)
                    current += len(photo_chunk)
                    self.progress_var.set(current)
                    self.status_var.set(f"İşleniyor: {category} sayfa {page_index}")
                    self.root.update_idletasks()

            desktop = Path.home() / "Desktop"
            default_name = default_output_filename(selected_categories)
            out_file = filedialog.asksaveasfilename(
                title="Birleştirilmiş Word kaydet",
                initialdir=str(desktop if desktop.exists() else Path.cwd()),
                initialfile=default_name,
                defaultextension=".docx",
                filetypes=[("Word", "*.docx")],
            )
            if not out_file:
                self.status_var.set("Kaydetme iptal edildi")
                return

            final_path = merge_docx_files(temp_outputs, Path(out_file))
            self.status_var.set(f"Tamamlandı: {final_path}")
            messagebox.showinfo("Başarılı", f"Dosya oluşturuldu:\n{final_path}")
        except Exception as exc:
            messagebox.showerror("Hata", str(exc))
            self.status_var.set(f"Hata: {exc}")

    def _on_exit(self):
        has_temp = TEMP_DIR.exists() and any(TEMP_DIR.iterdir())
        if has_temp:
            if messagebox.askyesno("Çıkış", "Tempfiles içeriği silinsin mi?"):
                self._clean_temp_dir()
        self.root.destroy()


def main():
    root = tk.Tk()
    WordMergeApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
