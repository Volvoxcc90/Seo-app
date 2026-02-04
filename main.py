# main.py
import os
import sys
import json
import re
import inspect
from pathlib import Path

# чтобы wb_fill гарантированно импортировался (и в exe тоже)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QPushButton, QFileDialog, QLineEdit,
    QVBoxLayout, QHBoxLayout, QGridLayout, QComboBox, QMessageBox,
    QProgressBar, QFrame, QCheckBox, QInputDialog
)
from PyQt5.QtCore import QThread, pyqtSignal, Qt

from wb_fill import fill_wb_template, DEFAULT_THEMES


APP_NAME = "Sunglasses SEO PRO"


# -------------------------------
# AppData paths
# -------------------------------
def app_root_dir() -> Path:
    base = Path(os.getenv("APPDATA", str(Path.home())))
    root = base / APP_NAME
    root.mkdir(parents=True, exist_ok=True)
    return root


def data_dir() -> Path:
    d = app_root_dir() / "data"
    d.mkdir(parents=True, exist_ok=True)
    return d


def settings_path() -> Path:
    return app_root_dir() / "settings.json"


def load_settings() -> dict:
    p = settings_path()
    if p.exists():
        try:
            return json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def save_settings(s: dict) -> None:
    settings_path().write_text(json.dumps(s, ensure_ascii=False, indent=2), encoding="utf-8")


def ensure_textfile(name: str, default_lines: list[str]) -> Path:
    p = data_dir() / name
    if not p.exists():
        p.write_text("\n".join(default_lines), encoding="utf-8")
    return p


def load_list_txt(name: str, default_lines: list[str]) -> list[str]:
    p = ensure_textfile(name, default_lines)
    items = [x.strip() for x in p.read_text(encoding="utf-8").splitlines() if x.strip()]
    # уникальность, сохранение порядка
    seen = set()
    out = []
    for it in items:
        if it not in seen:
            seen.add(it)
            out.append(it)
    return out


def append_to_txt(name: str, value: str) -> None:
    value = (value or "").strip()
    if not value:
        return
    p = data_dir() / name
    items = load_list_txt(name, [])
    if value not in items:
        with p.open("a", encoding="utf-8") as f:
            f.write("\n" + value)


def normalize_key(s: str) -> str:
    s = (s or "").strip().lower()
    s = s.replace("-", " ").replace("&", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def load_brands_ru_map() -> dict:
    p = data_dir() / "brands_ru.json"
    if p.exists():
        try:
            return json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def save_brands_ru_map(m: dict) -> None:
    p = data_dir() / "brands_ru.json"
    p.write_text(json.dumps(m, ensure_ascii=False, indent=2), encoding="utf-8")


# -------------------------------
# Worker thread
# -------------------------------
class Worker(QThread):
    progress = pyqtSignal(int)
    done = pyqtSignal(str, int)
    error = pyqtSignal(str)

    def __init__(self, args: dict):
        super().__init__()
        self.args = args

    def run(self):
        try:
            sig = inspect.signature(fill_wb_template)
            allowed = sig.parameters.keys()
            safe_args = {k: v for k, v in self.args.items() if k in allowed}
            if "progress_callback" in allowed:
                safe_args["progress_callback"] = lambda p: self.progress.emit(int(p))
            out_path, rows = fill_wb_template(**safe_args)
            self.done.emit(out_path, rows)
        except Exception as e:
            self.error.emit(str(e))


# -------------------------------
# UI
# -------------------------------
class App(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.setMinimumWidth(860)

        self.settings = load_settings()

        # seed data files
        self.brands = load_list_txt("brands.txt", ["Gucci", "Prada", "Miu Miu", "Ray-Ban", "Cazal"])
        self.shapes = load_list_txt("shapes.txt", ["авиаторы", "квадратные", "овальные", "кошачий глаз", "круглые"])
        self.lenses = load_list_txt("lenses.txt", ["UV400", "поляризационные", "фотохромные", "градиентные", "зеркальные"])
        ensure_textfile("brands_ru.json", ["{}"])  # пустой json

        self._build_ui()
        self.apply_theme(self.settings.get("theme", "Midnight"))

        # restore values
        self.theme_cb.setCurrentText(self.settings.get("theme", "Midnight"))
        self.brand_cb.setCurrentText(self.settings.get("brand", self.brands[0] if self.brands else ""))
        self.shape_cb.setCurrentText(self.settings.get("shape", self.shapes[0] if self.shapes else ""))
        self.lens_cb.setCurrentText(self.settings.get("lens", self.lenses[0] if self.lenses else ""))
        self.collection_cb.setCurrentText(self.settings.get("collection", "Весна–Лето 2026"))
        self.style_cb.setCurrentText(self.settings.get("style", "premium"))
        self.length_cb.setCurrentText(self.settings.get("length", "medium"))
        self.seo_cb.setCurrentText(self.settings.get("seo_level", "normal"))
        self.gender_cb.setCurrentText(self.settings.get("gender_mode", "Auto"))
        self.safe_chk.setChecked(bool(self.settings.get("wb_safe_mode", True)))
        self.strict_chk.setChecked(bool(self.settings.get("wb_strict", True)))

        self.input_xlsx = ""

    def _card(self) -> QFrame:
        c = QFrame()
        c.setObjectName("Card")
        c.setFrameShape(QFrame.NoFrame)
        c.setProperty("class", "card")
        return c

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(16, 16, 16, 16)
        root.setSpacing(12)

        # Header
        header = self._card()
        hl = QVBoxLayout(header)
        hl.setContentsMargins(16, 16, 16, 14)

        title = QLabel("🕶  Sunglasses SEO PRO")
        title.setObjectName("Title")
        subtitle = QLabel("Живые SEO-описания • Выпадающие списки • Прогресс • Темы • WB Safe/Strict • AUTO-пол")
        subtitle.setObjectName("Subtitle")
        hl.addWidget(title)
        hl.addWidget(subtitle)
        root.addWidget(header)

        # Theme row
        theme_card = self._card()
        tl = QHBoxLayout(theme_card)
        tl.setContentsMargins(16, 12, 16, 12)
        tl.setSpacing(10)

        tl.addWidget(QLabel("🎨 Тема"))
        self.theme_cb = QComboBox()
        self.theme_cb.addItems(list(DEFAULT_THEMES.keys()))
        self.theme_cb.currentTextChanged.connect(self.on_theme_change)
        self.theme_cb.setFixedWidth(190)
        tl.addWidget(self.theme_cb)

        tl.addSpacing(8)
        tl.addWidget(QLabel("📁 Справочники:"))
        self.data_path_lbl = QLineEdit(str(data_dir()))
        self.data_path_lbl.setReadOnly(True)
        tl.addWidget(self.data_path_lbl, 1)

        self.open_data_btn = QPushButton("Папка")
        self.open_data_btn.clicked.connect(self.open_data_folder)
        tl.addWidget(self.open_data_btn)

        root.addWidget(theme_card)

        # File
        file_card = self._card()
        fl = QHBoxLayout(file_card)
        fl.setContentsMargins(16, 12, 16, 12)
        fl.setSpacing(10)

        self.pick_btn = QPushButton("📄 Загрузить XLSX")
        self.pick_btn.clicked.connect(self.pick_xlsx)
        fl.addWidget(self.pick_btn)

        self.file_lbl = QLabel("Файл не выбран")
        self.file_lbl.setObjectName("Muted")
        self.file_lbl.setTextInteractionFlags(Qt.TextSelectableByMouse)
        fl.addWidget(self.file_lbl, 1)

        root.addWidget(file_card)

        # Main inputs
        form_card = self._card()
        gl = QGridLayout(form_card)
        gl.setContentsMargins(16, 14, 16, 14)
        gl.setHorizontalSpacing(10)
        gl.setVerticalSpacing(10)

        # Brand row
        gl.addWidget(QLabel("Бренд"), 0, 0)
        self.brand_cb = QComboBox()
        self.brand_cb.setEditable(True)
        self.brand_cb.addItems(self.brands)
        gl.addWidget(self.brand_cb, 0, 1)
        bplus = QPushButton("+")
        bplus.setFixedWidth(48)
        bplus.clicked.connect(self.add_brand)
        gl.addWidget(bplus, 0, 2)

        # Shape
        gl.addWidget(QLabel("Форма оправы"), 1, 0)
        self.shape_cb = QComboBox()
        self.shape_cb.setEditable(True)
        self.shape_cb.addItems(self.shapes)
        gl.addWidget(self.shape_cb, 1, 1)
        splus = QPushButton("+")
        splus.setFixedWidth(48)
        splus.clicked.connect(self.add_shape)
        gl.addWidget(splus, 1, 2)

        # Lens
        gl.addWidget(QLabel("Линзы"), 2, 0)
        self.lens_cb = QComboBox()
        self.lens_cb.setEditable(True)
        self.lens_cb.addItems(self.lenses)
        gl.addWidget(self.lens_cb, 2, 1)
        lplus = QPushButton("+")
        lplus.setFixedWidth(48)
        lplus.clicked.connect(self.add_lens)
        gl.addWidget(lplus, 2, 2)

        # Collection
        gl.addWidget(QLabel("Коллекция"), 3, 0)
        self.collection_cb = QComboBox()
        self.collection_cb.setEditable(True)
        self.collection_cb.addItems(["Весна–Лето 2026", "Весна–Лето 2025–2026"])
        gl.addWidget(self.collection_cb, 3, 1, 1, 2)

        # Controls row: SEO/length/style
        gl.addWidget(QLabel("SEO-плотность"), 4, 0)
        self.seo_cb = QComboBox()
        self.seo_cb.addItems(["low", "normal", "high"])
        gl.addWidget(self.seo_cb, 4, 1)

        gl.addWidget(QLabel("Длина"), 4, 2)
        self.length_cb = QComboBox()
        self.length_cb.addItems(["short", "medium", "long"])
        gl.addWidget(self.length_cb, 4, 3)

        gl.addWidget(QLabel("Стиль"), 4, 4)
        self.style_cb = QComboBox()
        self.style_cb.addItems(["neutral", "premium", "mass", "social"])
        gl.addWidget(self.style_cb, 4, 5)

        # AUTO gender + safe/strict
        gl.addWidget(QLabel("AUTO-пол"), 5, 0)
        self.gender_cb = QComboBox()
        self.gender_cb.addItems(["Auto", "Женские", "Мужские", "Унисекс"])
        gl.addWidget(self.gender_cb, 5, 1)

        self.safe_chk = QCheckBox("WB Safe Mode (заменяет риск-слова)")
        gl.addWidget(self.safe_chk, 5, 2, 1, 2)

        self.strict_chk = QCheckBox("WB Strict (убирает обещания/абсолюты)")
        gl.addWidget(self.strict_chk, 5, 4, 1, 2)

        root.addWidget(form_card)

        # Progress + Run
        bottom = self._card()
        bl = QHBoxLayout(bottom)
        bl.setContentsMargins(16, 12, 16, 12)
        bl.setSpacing(12)

        self.progress = QProgressBar()
        self.progress.setValue(0)
        bl.addWidget(self.progress, 1)

        self.run_btn = QPushButton("🚀  СГЕНЕРИРОВАТЬ")
        self.run_btn.setFixedHeight(44)
        self.run_btn.clicked.connect(self.run)
        bl.addWidget(self.run_btn)

        root.addWidget(bottom)

    # ---------- themes ----------
    def apply_theme(self, name: str):
        qss = DEFAULT_THEMES.get(name, DEFAULT_THEMES["Midnight"])
        self.setStyleSheet(qss)

    def on_theme_change(self, name: str):
        self.apply_theme(name)
        self.settings["theme"] = name
        save_settings(self.settings)

    # ---------- data ----------
    def open_data_folder(self):
        try:
            os.startfile(str(data_dir()))
        except Exception:
            QMessageBox.information(self, "data", str(data_dir()))

    # ---------- add items ----------
    def add_brand(self):
        lat, ok = QInputDialog.getText(self, "Добавить бренд (латиницей)", "Например: Miu Miu")
        if not ok:
            return
        lat = (lat or "").strip()
        if not lat:
            return

        # спросим кириллицу, но уже с подсказкой (простая транслитерация)
        guess = self._guess_ru(lat)
        ru, ok2 = QInputDialog.getText(self, "Кириллица для наименования", "Как писать в названии (RU):", text=guess)
        if not ok2:
            return
        ru = (ru or "").strip()
        if not ru:
            ru = guess

        append_to_txt("brands.txt", lat)
        self.brands = load_list_txt("brands.txt", [])
        self.brand_cb.clear()
        self.brand_cb.addItems(self.brands)
        self.brand_cb.setCurrentText(lat)

        # сохранить mapping lat->ru
        m = load_brands_ru_map()
        m[normalize_key(lat)] = ru
        save_brands_ru_map(m)

    def add_shape(self):
        val, ok = QInputDialog.getText(self, "Добавить форму оправы", "Например: квадратные")
        if not ok:
            return
        val = (val or "").strip()
        if not val:
            return
        append_to_txt("shapes.txt", val)
        self.shapes = load_list_txt("shapes.txt", [])
        self.shape_cb.clear()
        self.shape_cb.addItems(self.shapes)
        self.shape_cb.setCurrentText(val)

    def add_lens(self):
        val, ok = QInputDialog.getText(self, "Добавить линзы/особенности", "Например: UV400, фотохромные")
        if not ok:
            return
        val = (val or "").strip()
        if not val:
            return
        append_to_txt("lenses.txt", val)
        self.lenses = load_list_txt("lenses.txt", [])
        self.lens_cb.clear()
        self.lens_cb.addItems(self.lenses)
        self.lens_cb.setCurrentText(val)

    def _guess_ru(self, brand: str) -> str:
        # простая транслитерация как подсказка
        t = brand.strip().lower()
        t = t.replace("-", " ")
        t = re.sub(r"\s+", " ", t).strip()
        rules = [
            ("sch","ш"),("sh","ш"),("ch","ч"),("ya","я"),("yu","ю"),("yo","ё"),
            ("kh","х"),("ts","ц"),("ph","ф"),("th","т"),
            ("a","а"),("b","б"),("c","к"),("d","д"),("e","е"),("f","ф"),
            ("g","г"),("h","х"),("i","и"),("j","дж"),("k","к"),("l","л"),
            ("m","м"),("n","н"),("o","о"),("p","п"),("q","к"),("r","р"),
            ("s","с"),("t","т"),("u","у"),("v","в"),("w","в"),("x","кс"),
            ("y","и"),("z","з"),
        ]
        out = []
        for w in t.split():
            ww = w
            for a, b in rules:
                ww = ww.replace(a, b)
            out.append(ww.capitalize())
        return " ".join(out) if out else brand

    # ---------- xlsx ----------
    def pick_xlsx(self):
        fp, _ = QFileDialog.getOpenFileName(self, "Выбрать XLSX", "", "Excel (*.xlsx)")
        if fp:
            self.input_xlsx = fp
            self.file_lbl.setText(fp)
            self.file_lbl.setObjectName("")
            self.file_lbl.style().unpolish(self.file_lbl)
            self.file_lbl.style().polish(self.file_lbl)

    def run(self):
        if not self.input_xlsx:
            QMessageBox.warning(self, "Ошибка", "Выбери XLSX файл")
            return

        # save last inputs
        self.settings.update({
            "brand": self.brand_cb.currentText(),
            "shape": self.shape_cb.currentText(),
            "lens": self.lens_cb.currentText(),
            "collection": self.collection_cb.currentText(),
            "style": self.style_cb.currentText(),
            "length": self.length_cb.currentText(),
            "seo_level": self.seo_cb.currentText(),
            "gender_mode": self.gender_cb.currentText(),
            "wb_safe_mode": self.safe_chk.isChecked(),
            "wb_strict": self.strict_chk.isChecked(),
            "theme": self.theme_cb.currentText(),
        })
        save_settings(self.settings)

        args = dict(
            input_xlsx=self.input_xlsx,
            brand_lat=self.brand_cb.currentText().strip(),
            shape=self.shape_cb.currentText().strip(),
            lens=self.lens_cb.currentText().strip(),
            collection=self.collection_cb.currentText().strip(),
            style=self.style_cb.currentText().strip(),
            desc_length=self.length_cb.currentText().strip(),
            seo_level=self.seo_cb.currentText().strip(),
            gender_mode=self.gender_cb.currentText().strip(),
            wb_safe_mode=self.safe_chk.isChecked(),
            wb_strict=self.strict_chk.isChecked(),
            data_dir=str(data_dir()),
        )

        self.progress.setValue(0)
        self.run_btn.setEnabled(False)

        self.worker = Worker(args)
        self.worker.progress.connect(self.progress.setValue)
        self.worker.done.connect(self.on_done)
        self.worker.error.connect(self.on_error)
        self.worker.start()

    def on_done(self, out_path: str, rows: int):
        self.progress.setValue(100)
        self.run_btn.setEnabled(True)
        QMessageBox.information(self, "Готово", f"Готовый файл:\n{out_path}\n\nСтрок обработано: {rows}")

    def on_error(self, msg: str):
        self.run_btn.setEnabled(True)
        QMessageBox.critical(self, "Ошибка", msg)


def main():
    app = QApplication(sys.argv)
    w = App()
    w.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
