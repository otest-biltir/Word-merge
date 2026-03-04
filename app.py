from __future__ import annotations

import shutil
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from docx.shared import Mm
from docxtpl import DocxTemplate, InlineImage
from docxcompose.composer import Composer
from docx import Document

BASE_DIR = Path(__file__).resolve().parent
TEMPLATES_DIR = BASE_DIR / "templates"
TEMP_DIR = BASE_DIR / "tempfiles"

CATEGORIES = [
    ("pre", "Pre photos", "pre_photos.docx"),
    ("post", "Post photos", "post_photos.docx"),
    ("teardown", "Teardown photos", "teardown_photos.docx"),
    ("handle_side_cover", "Handle side cover", "handle_side_cover.docx"),
]

SLOT_COUNT = 6
SLOT_WIDTH_MM = 55
SLOT_HEIGHT_MM = 45


class WordMergeApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Word Fotoğraf Birleştirici")
        self.root.geometry("900x620")

        TEMP_DIR.mkdir(exist_ok=True)
        TEMPLATES_DIR.mkdir(exist_ok=True)

        self.selected_files: dict[str, list[Path]] = {key: [] for key, _, _ in CATEGORIES}
        self.count_vars: dict[str, tk.StringVar] = {}

        self.progress_var = tk.DoubleVar(value=0)
        self.progress_total = 1
        self.progress_done = 0

        self._prompt_cleanup_on_start()
        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_ui(self) -> None:
        style = ttk.Style()
        if "clam" in style.theme_names():
            style.theme_use("clam")

        container = ttk.Frame(self.root, padding=16)
        container.pack(fill="both", expand=True)

        title = ttk.Label(
            container,
            text="Fotoğraf Yükle ve Word Birleştir",
            font=("Segoe UI", 18, "bold"),
        )
        title.pack(anchor="w", pady=(0, 8))

        subtitle = ttk.Label(
            container,
            text=(
                "Template dosyalarını templates klasörüne koyun. Her template içinde "
                "{{IMG_1}} ... {{IMG_6}} yer tutucuları olmalıdır."
            ),
            font=("Segoe UI", 10),
        )
        subtitle.pack(anchor="w", pady=(0, 16))

        cards = ttk.Frame(container)
        cards.pack(fill="x")

        for i, (key, label, template_name) in enumerate(CATEGORIES):
            card = ttk.LabelFrame(cards, text=label, padding=12)
            card.grid(row=i // 2, column=i % 2, padx=8, pady=8, sticky="nsew")

            cards.columnconfigure(i % 2, weight=1)
            cards.rowconfigure(i // 2, weight=1)

            count_var = tk.StringVar(value="0 fotoğraf seçildi")
            self.count_vars[key] = count_var

            ttk.Label(card, text=f"Template: templates/{template_name}").pack(anchor="w")
            ttk.Label(card, textvariable=count_var, font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(8, 10))

            ttk.Button(
                card,
                text="Fotoğraf Ekle",
                command=lambda k=key: self._add_photos(k),
            ).pack(side="left", padx=(0, 8))

            ttk.Button(
                card,
                text="Temizle",
                command=lambda k=key: self._clear_category(k),
            ).pack(side="left")

        progress_frame = ttk.Frame(container)
        progress_frame.pack(fill="x", pady=(20, 10))

        ttk.Label(progress_frame, text="Yükleme / Birleştirme Durumu").pack(anchor="w")

        self.progress = ttk.Progressbar(
            progress_frame,
            maximum=100,
            variable=self.progress_var,
            mode="determinate",
        )
        self.progress.pack(fill="x", pady=(6, 0))

        self.progress_text = tk.StringVar(value="Hazır")
        ttk.Label(progress_frame, textvariable=self.progress_text).pack(anchor="w", pady=(4, 0))

        actions = ttk.Frame(container)
        actions.pack(fill="x", pady=(16, 0))

        ttk.Button(actions, text="Word Dosyalarını Birleştir", command=self._merge).pack(side="left")
        ttk.Button(actions, text="Tempfiles Klasörünü Temizle", command=self._clear_tempfiles).pack(side="left", padx=8)

    def _set_progress(self, done: int | None = None, total: int | None = None, status: str | None = None) -> None:
        if total is not None:
            self.progress_total = max(1, total)
        if done is not None:
            self.progress_done = max(0, done)

        value = (self.progress_done / self.progress_total) * 100
        self.progress_var.set(value)

        if status:
            self.progress_text.set(status)

        self.root.update_idletasks()

    def _prompt_cleanup_on_start(self) -> None:
        if any(TEMP_DIR.iterdir()):
            delete = messagebox.askyesno(
                "Tempfiles bulundu",
                "Var olan tempfiles içeriği bulundu. Uygulama başlamadan temizlensin mi?",
            )
            if delete:
                self._clear_tempfiles(show_message=False)

    def _copy_to_temp(self, category: str, source: Path, index: int) -> Path:
        category_dir = TEMP_DIR / category
        category_dir.mkdir(parents=True, exist_ok=True)

        target = category_dir / f"{index:03d}_{source.name}"
        shutil.copy2(source, target)
        return target

    def _add_photos(self, category: str) -> None:
        files = filedialog.askopenfilenames(
            title="Fotoğraf seç",
            filetypes=[("Images", "*.png *.jpg *.jpeg *.bmp *.webp *.tif *.tiff")],
        )
        if not files:
            return

        existing = len(self.selected_files[category])
        total_steps = len(files)
        self._set_progress(done=0, total=total_steps, status="Fotoğraflar yükleniyor...")

        for step, file in enumerate(files, start=1):
            copied = self._copy_to_temp(category, Path(file), existing + step)
            self.selected_files[category].append(copied)
            self._set_progress(done=step, status=f"{step}/{total_steps} fotoğraf yüklendi")

        self._update_counts()

    def _clear_category(self, category: str) -> None:
        self.selected_files[category] = []
        category_dir = TEMP_DIR / category
        if category_dir.exists():
            shutil.rmtree(category_dir, ignore_errors=True)
        self._update_counts()

    def _update_counts(self) -> None:
        for key, _, _ in CATEGORIES:
            self.count_vars[key].set(f"{len(self.selected_files[key])} fotoğraf seçildi")

    def _default_output_name(self) -> str:
        used = [key for key, _, _ in CATEGORIES if self.selected_files[key]]
        return ("_".join(used) if used else "merged") + ".docx"

    def _render_chunk(self, template_path: Path, images: list[Path], out_path: Path) -> None:
        doc = DocxTemplate(str(template_path))

        context = {}
        for i in range(SLOT_COUNT):
            token = f"IMG_{i + 1}"
            if i < len(images):
                context[token] = InlineImage(
                    doc,
                    str(images[i]),
                    width=Mm(SLOT_WIDTH_MM),
                    height=Mm(SLOT_HEIGHT_MM),
                )
            else:
                context[token] = ""

        doc.render(context)
        doc.save(str(out_path))

    def _merge(self) -> None:
        active_categories = [(k, t) for k, _, t in CATEGORIES if self.selected_files[k]]
        if not active_categories:
            messagebox.showwarning("Uyarı", "Lütfen en az bir kategoriye fotoğraf ekleyin.")
            return

        desktop = Path.home() / "Desktop"
        default_name = self._default_output_name()
        output = filedialog.asksaveasfilename(
            title="Birleşik Word dosyasını kaydet",
            defaultextension=".docx",
            initialdir=str(desktop if desktop.exists() else BASE_DIR),
            initialfile=default_name,
            filetypes=[("Word Document", "*.docx")],
        )
        if not output:
            return

        render_dir = TEMP_DIR / "rendered"
        render_dir.mkdir(exist_ok=True)

        rendered_files: list[Path] = []
        total_chunks = 0
        for category, template_name in active_categories:
            chunk_count = (len(self.selected_files[category]) + SLOT_COUNT - 1) // SLOT_COUNT
            total_chunks += chunk_count

        self._set_progress(done=0, total=total_chunks + 1, status="Word sayfaları hazırlanıyor...")

        step = 0
        for category, template_name in active_categories:
            template_path = TEMPLATES_DIR / template_name
            if not template_path.exists():
                messagebox.showerror(
                    "Template bulunamadı",
                    f"Template eksik: {template_path}\nLütfen dosyayı templates klasörüne ekleyin.",
                )
                return

            images = self.selected_files[category]
            chunks = [images[i : i + SLOT_COUNT] for i in range(0, len(images), SLOT_COUNT)]
            for idx, chunk in enumerate(chunks, start=1):
                out_path = render_dir / f"{category}_{idx:03d}.docx"
                self._render_chunk(template_path, chunk, out_path)
                rendered_files.append(out_path)
                step += 1
                self._set_progress(done=step, status=f"Sayfa hazırlanıyor: {category} #{idx}")

        base_doc = Document(str(rendered_files[0]))
        composer = Composer(base_doc)
        for doc_path in rendered_files[1:]:
            composer.append(Document(str(doc_path)))

        composer.save(output)
        self._set_progress(done=total_chunks + 1, status=f"Tamamlandı: {output}")

        messagebox.showinfo("Başarılı", f"Birleştirme tamamlandı.\nDosya: {output}")

    def _clear_tempfiles(self, show_message: bool = True) -> None:
        for item in TEMP_DIR.iterdir():
            if item.is_dir():
                shutil.rmtree(item, ignore_errors=True)
            else:
                item.unlink(missing_ok=True)
        if show_message:
            messagebox.showinfo("Tamam", "Tempfiles klasörü temizlendi.")

    def _on_close(self) -> None:
        should_delete = messagebox.askyesno(
            "Çıkış",
            "Uygulamadan çıkarken tempfiles klasörü temizlensin mi?",
        )
        if should_delete:
            self._clear_tempfiles(show_message=False)
        self.root.destroy()


def main() -> None:
    root = tk.Tk()
    app = WordMergeApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
