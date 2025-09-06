import os
import numpy as np
import pandas as pd
import dearpygui.dearpygui as dpg

SAMPLE_PATH = "sample_data.xlsx"


def read_three_cols(path):
    """Читает Excel и возвращает датафрейм колонками ['param','x','y']"""
    df = pd.read_excel(path, engine="openpyxl")
    if df.shape[1] >= 3:
        df = df.iloc[:, :3].copy()
        df.columns = ["param", "x", "y"]
    else:
        df = df.iloc[:, :2].copy()
        df.columns = ["x", "y"]
        df["param"] = df.index.astype(str)
        df = df[["param", "x", "y"]]
    df["x"] = pd.to_numeric(df["x"], errors="coerce")
    df["y"] = pd.to_numeric(df["y"], errors="coerce")
    df = df.dropna(subset=["x", "y"]).reset_index(drop=True)
    return df


class DPGBasedApp:
    def __init__(self):
        self.plot_tag = "main_plot"
        self.x_axis_tag = "x_axis"
        self.y_axis_tag = "y_axis"
        self.data_series_tag = "data_series"
        self.selected_series_tag = "selected_series"
        self.list_tag = "coords_list"

        self.df = None
        self.xmin = self.xmax = self.ymin = self.ymax = None

    def start(self):
        dpg.create_context()
        with dpg.font_registry():
            with dpg.font("C:/test/src/font/DejaVuSans.ttf", 14, tag="rus_font"):
                dpg.add_font_range_hint(dpg.mvFontRangeHint_Default)
                dpg.add_font_range_hint(dpg.mvFontRangeHint_Cyrillic)
                
        dpg.bind_font("rus_font")
          
        dpg.create_viewport(title="Excel → График", width=1000, height=650)

        # UI
        with dpg.window(label="Главное окно", width=1000, height=650):
            # Кнопки
            with dpg.group(horizontal=True):
                dpg.add_button(label="Загрузить sample_data.xlsx", callback=self.load_sample)
                # Кнопка без функционала пока
                dpg.add_button(label="Загрузить другой файл...", callback=lambda s, a: None)

            dpg.add_separator()

            # Основная область: график слева, список справа
            with dpg.group(horizontal=True):
                with dpg.child_window(width=720, height=540):
                    with dpg.plot(label="График", height=-1, width=-1, tag=self.plot_tag):
                        dpg.add_plot_legend()
                        dpg.add_plot_axis(dpg.mvXAxis, label="X", tag=self.x_axis_tag)
                        with dpg.plot_axis(dpg.mvYAxis, label="Y", tag=self.y_axis_tag):
                            # серии, указываем parent на Y-ось
                            dpg.add_scatter_series([], [], label="данные", tag=self.data_series_tag, parent=self.y_axis_tag)
                            dpg.add_scatter_series([], [], label="выделение", tag=self.selected_series_tag, parent=self.y_axis_tag)
                with dpg.child_window(width=250, height=540):
                    dpg.add_text("Найденная ближайшая точка:")
                    dpg.add_listbox(items=[], num_items=10, tag=self.list_tag)

        # Регистрируем обработчики в handler_registry (требование DearPyGui)
        with dpg.handler_registry():
            dpg.add_mouse_release_handler(callback=self.on_mouse_release)

        dpg.setup_dearpygui()
        dpg.show_viewport()
        dpg.start_dearpygui()
        dpg.destroy_context()

    # загрузка образца
    def load_sample(self, sender, app_data):
        path = os.path.join(os.getcwd(), SAMPLE_PATH)
        if not os.path.exists(path):
            self._create_sample()
        try:
            df = read_three_cols(path)
            self._apply_dataframe(df)
        except Exception:
            pass

    def _apply_dataframe(self, df):
        """Сохраняем df, считаем границы и обновляем график/listbox."""
        self.df = df.copy().reset_index(drop=True)
        xs = self.df["x"].tolist()
        ys = self.df["y"].tolist()

        if xs:
            xminv, xmaxv = min(xs), max(xs)
            padx = (xmaxv - xminv) * 0.05 if xmaxv != xminv else 1.0
            self.xmin, self.xmax = xminv - padx, xmaxv + padx
        else:
            self.xmin, self.xmax = -1.0, 1.0

        if ys:
            yminv, ymaxv = min(ys), max(ys)
            pady = (ymaxv - yminv) * 0.05 if ymaxv != yminv else 1.0
            self.ymin, self.ymax = yminv - pady, ymaxv + pady
        else:
            self.ymin, self.ymax = -1.0, 1.0

        dpg.set_value(self.data_series_tag, [xs, ys])

        # лимиты осей (на случай несовместимости)
        try:
            dpg.set_axis_limits(self.x_axis_tag, float(self.xmin), float(self.xmax))
            dpg.set_axis_limits(self.y_axis_tag, float(self.ymin), float(self.ymax))
        except Exception:
            pass

        dpg.set_value(self.selected_series_tag, [[], []])

        # обновить список всех точек (кратко)
        items = [f"{r.param} — x={r.x} y={r.y}" for r in self.df.itertuples()]
        dpg.configure_item(self.list_tag, items=items)

    def on_mouse_release(self, sender, app_data):
        """Если курсор над графиком то находим ближайшую точку."""
        if self.df is None:
            return

        try:
            pos = dpg.get_plot_mouse_pos()
        except Exception:
            pos = None

        if not pos:
            return

        x_mouse, y_mouse = pos
        if x_mouse is None or y_mouse is None:
            return

        xs = np.array(self.df["x"], dtype=float)
        ys = np.array(self.df["y"], dtype=float)
        d2 = (xs - x_mouse) ** 2 + (ys - y_mouse) ** 2
        if d2.size == 0:
            return
        idx = int(np.argmin(d2))
        row = self.df.iloc[idx]

        # обновить выделение на графике (одна точка)
        dpg.set_value(self.selected_series_tag, [[row.x], [row.y]])

        # показать только найденную точку в списке
        item_text = f"{row.param} — x={row.x} y={row.y}"
        dpg.configure_item(self.list_tag, items=[item_text])

    def _create_sample(self):
        df = pd.DataFrame({
            "param": [f"p{i}" for i in range(1, 21)],
            "x": np.linspace(0, 10, 20),
            "y": np.sin(np.linspace(0, 10, 20)) * 3 + np.linspace(-1, 1, 20)
        })
        df.to_excel(SAMPLE_PATH, index=False)


if __name__ == "__main__":
    app = DPGBasedApp()
    app.start()