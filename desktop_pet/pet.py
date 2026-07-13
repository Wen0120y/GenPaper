from __future__ import annotations

import argparse
import json
import math
import sys
import tkinter as tk
import webbrowser
from pathlib import Path
from tkinter import Menu

from PIL import Image, ImageChops, ImageDraw, ImageStat, ImageTk


ROOT = Path(__file__).resolve().parent
SPRITESHEET = ROOT / "spritesheet.webp"
VALIDATION = ROOT / "validation.json"
CUSTOM_ACTIONS = ROOT / "custom_actions.webp"
TRANSPARENT_COLOR = "#010203"
ACTA_URL = "https://actaresearch.streamlit.app/"

ACTION_STATES = {
    "idle": "idle",
    "rest": "waiting",
    "paper": "running",
    "experiment": "review",
    "think": "thinking",
    "wave": "waving",
    "jump": "jumping",
}

TIMED_ACTIONS = {
    "rest": "休息",
    "paper": "论文",
    "experiment": "实验",
}

TIMES = {
    "5 分钟": 5 * 60,
    "30 分钟": 30 * 60,
    "1 小时": 60 * 60,
    "2 小时": 2 * 60 * 60,
}


SIZE_OPTIONS = {
    "60%": 0.6,
    "80%": 0.8,
    "100%": 1.0,
    "120%": 1.2,
}


class CodexDesktopPet:
    def __init__(self, scale: float) -> None:
        self.root = tk.Tk()
        self.root.overrideredirect(True)
        self.root.attributes("-topmost", True)
        self.root.configure(bg=TRANSPARENT_COLOR)
        self.root.wm_attributes("-transparentcolor", TRANSPARENT_COLOR)

        self.scale = scale
        self.frames = self._load_frames(scale)
        self.state = "idle"
        self.frame_index = 0
        self.timer_job: str | None = None
        self.click_job: str | None = None
        self.one_shot_state: str | None = None
        self.pointer_down = False
        self.dragging = False
        self.press_button = 0
        self.drag_offset_x = 0
        self.drag_offset_y = 0
        self.press_root_x = 0
        self.press_root_y = 0
        self.last_left_release_time = 0
        self.last_left_release_x = -10_000
        self.last_left_release_y = -10_000
        self.last_pointer_x = 0
        self.drag_state = "running-right"
        self.return_state = "idle"

        self.label = tk.Label(
            self.root,
            bg=TRANSPARENT_COLOR,
            bd=0,
            highlightthickness=0,
            cursor="hand2",
        )
        self.label.pack()
        self.label.configure(image=self.frames["idle"][0])

        for widget in (self.root, self.label):
            widget.bind("<ButtonPress-1>", self._start_drag)
            widget.bind("<B1-Motion>", self._drag)
            widget.bind("<ButtonRelease-1>", self._stop_drag)
            widget.bind("<ButtonPress-3>", self._right_press)
            widget.bind("<B3-Motion>", self._drag)
            widget.bind("<ButtonRelease-3>", self._right_release)

        self.menu = self._build_menu()
        self._place_initially()
        self._tick()

    def run(self) -> None:
        self.root.mainloop()

    def _load_frames(self, scale: float) -> dict[str, list[ImageTk.PhotoImage]]:
        if not SPRITESHEET.exists() or not VALIDATION.exists():
            raise FileNotFoundError("需要 spritesheet.webp 和 validation.json 与 pet.py 放在同一目录。")

        with VALIDATION.open("r", encoding="utf-8") as file:
            metadata = json.load(file)

        sheet = Image.open(SPRITESHEET).convert("RGBA")
        cell_width = sheet.width // metadata["columns"]
        cell_height = sheet.height // metadata["rows"]

        grouped_cells: dict[str, list[tuple[int, int]]] = {}
        for cell in metadata["cells"]:
            if cell.get("used"):
                grouped_cells.setdefault(cell["state"], []).append((cell["row"], cell["column"]))

        pil_frames: dict[str, list[Image.Image]] = {}
        for state, cells in grouped_cells.items():
            state_frames = []
            for row, column in sorted(cells):
                left = column * cell_width
                top = row * cell_height
                frame = sheet.crop((left, top, left + cell_width, top + cell_height))
                frame = self._remove_halo(frame)
                if scale != 1:
                    size = (round(cell_width * scale), round(cell_height * scale))
                    frame = frame.resize(size, Image.Resampling.LANCZOS)
                    frame = self._remove_halo(frame)
                state_frames.append(frame)
            pil_frames[state] = self._hatch_frames(state, state_frames)

        if "idle" not in pil_frames:
            raise ValueError("validation.json 中没有可用的 idle 帧。")
        if "thinking" not in pil_frames:
            pil_frames["thinking"] = self._hatch_thinking(pil_frames["idle"])
        pil_frames.update(self._load_custom_action_frames(scale))

        return {
            state: [ImageTk.PhotoImage(frame) for frame in state_frames]
            for state, state_frames in pil_frames.items()
        }

    def _load_custom_action_frames(self, scale: float) -> dict[str, list[Image.Image]]:
        if not CUSTOM_ACTIONS.exists():
            return {}

        sheet = Image.open(CUSTOM_ACTIONS).convert("RGBA")
        columns = 10
        rows = 3
        cell_width = sheet.width // columns
        cell_height = sheet.height // rows
        states = ("waiting", "thinking", "running")
        actions: dict[str, list[Image.Image]] = {}

        for row, state in enumerate(states):
            frames = []
            for column in range(columns):
                left = column * cell_width
                top = row * cell_height
                frame = sheet.crop((left, top, left + cell_width, top + cell_height))
                frame = self._remove_halo(frame)
                if scale != 1:
                    size = (round(cell_width * scale), round(cell_height * scale))
                    frame = frame.resize(size, Image.Resampling.LANCZOS)
                    frame = self._remove_halo(frame)
                frames.append(frame)
            actions[state] = frames

        return actions

    def _remove_halo(self, frame: Image.Image) -> Image.Image:
        frame = frame.convert("RGBA")
        red, green, blue, alpha = frame.split()
        alpha = alpha.point(lambda value: 255 if value >= 64 else 0)
        return Image.merge("RGBA", (red, green, blue, alpha))

    def _hatch_frames(self, state: str, frames: list[Image.Image]) -> list[Image.Image]:
        if state == "failed":
            return frames
        if state == "waiting":
            return self._hatch_rest(frames)
        if state == "running":
            return self._hatch_paper(frames)
        if state == "review":
            return self._hatch_lab(frames)
        if state == "thinking":
            return self._hatch_thinking(frames)

        if len(frames) < 2 or not self._is_mostly_static(frames):
            return frames

        base = max(frames, key=self._visible_height)
        return self._synthesize_motion(base, count=8, bob=1.5, squash=0.01, tilt=0.5)

    def _hatch_thinking(self, frames: list[Image.Image]) -> list[Image.Image]:
        base = frames[0]
        result = []
        for index in range(10):
            frame = base.copy()
            self._draw_idea_bulb(frame, index)
            result.append(self._remove_halo(frame))
        return result

    def _hatch_rest(self, frames: list[Image.Image]) -> list[Image.Image]:
        base = frames[3] if len(frames) > 3 else frames[0]
        result = []
        for index in range(10):
            phase = math.sin(index / 10 * math.tau)
            frame = self._transform_subject(base, dx=0, dy=round(-2 * max(0, phase)), angle=-2.8 * max(0, phase))
            self._draw_coffee_steam(frame, index)
            result.append(self._remove_halo(frame))
        return result

    def _hatch_paper(self, frames: list[Image.Image]) -> list[Image.Image]:
        base = max(frames, key=self._visible_height)
        result = []
        for index in range(10):
            frame = base.copy()
            self._draw_coding_head_motion(frame, index)
            self._draw_typing_hands(frame, index)
            if 4 <= index <= 7:
                self._draw_closed_eye_smile(frame)
            self._typing_ticks(frame, index)
            result.append(self._remove_halo(frame))
        return result

    def _hatch_lab(self, frames: list[Image.Image]) -> list[Image.Image]:
        base = max(frames, key=self._visible_height)
        result = []
        for index in range(12):
            frame = base.copy()
            self._draw_lab_focus(frame, index)
            self._liquid_pulse(frame, index)
            self._beaker_bubbles(frame, index)
            result.append(self._remove_halo(frame))
        return result

    def _visible_height(self, image: Image.Image) -> int:
        bbox = image.getbbox()
        return bbox[3] - bbox[1] if bbox else 0

    def _box(self, image: Image.Image, relative_box: tuple[float, float, float, float]) -> tuple[int, int, int, int]:
        width, height = image.size
        left, top, right, bottom = relative_box
        return (
            round(width * left),
            round(height * top),
            round(width * right),
            round(height * bottom),
        )

    def _draw_idea_bulb(self, image: Image.Image, index: int) -> None:
        draw = ImageDraw.Draw(image, "RGBA")
        glow_levels = [0, 28, 48, 70, 96, 118, 165, 118, 58, 20]
        glow = glow_levels[index % len(glow_levels)]
        pulse = index in {5, 6, 7}
        cx = round(image.width * 0.66)
        cy = round(image.height * 0.17)
        radius = round(image.width * (0.078 if index in {6, 7} else 0.058))
        draw.ellipse((cx - radius - 6, cy - radius - 6, cx + radius + 6, cy + radius + 6), fill=(255, 225, 80, glow))
        bulb_alpha = 235 if glow else 80
        draw.ellipse((cx - radius, cy - radius, cx + radius, cy + radius), fill=(255, 238, 92, bulb_alpha), outline=(80, 58, 30, 210), width=2)
        stem_w = round(radius * 0.72)
        draw.rounded_rectangle((cx - stem_w // 2, cy + radius - 1, cx + stem_w // 2, cy + radius + 8), radius=2, fill=(238, 218, 145, 245), outline=(80, 58, 30, 230), width=1)
        if pulse:
            for dx, dy in ((-18, -14), (0, -22), (18, -14), (22, 4)):
                draw.line((cx + dx * 0.7, cy + dy * 0.7, cx + dx, cy + dy), fill=(255, 235, 110, 210), width=2)

    def _draw_coffee_steam(self, image: Image.Image, index: int) -> None:
        self._clear_steam_pixels(image)
        if index >= 6:
            return

        draw = ImageDraw.Draw(image, "RGBA")
        x = round(image.width * 0.48)
        y = round(image.height * 0.48)
        for steam in range(2):
            if (index + steam) % 3 != 0:
                sx = x + steam * 8
                draw.arc((sx, y - 12, sx + 6, y - 2), start=260, end=80, fill=(180, 180, 180, 120), width=1)

    def _clear_steam_pixels(self, image: Image.Image) -> None:
        left, top, right, bottom = self._box(image, (0.0, 0.25, 0.34, 0.60))
        pixels = image.load()
        for y in range(top, bottom):
            for x in range(left, right):
                red, green, blue, alpha = pixels[x, y]
                low_saturation = max(red, green, blue) - min(red, green, blue) < 60
                steam_tone = 105 <= red <= 230 and green > 95 and blue > 80
                if alpha > 0 and low_saturation and steam_tone:
                    pixels[x, y] = (0, 0, 0, 0)

    def _draw_coding_head_motion(self, image: Image.Image, index: int) -> None:
        draw = ImageDraw.Draw(image, "RGBA")
        if index in {1, 3, 5, 7}:
            side = -1 if index in {1, 5} else 1
            x = round(image.width * (0.26 if side < 0 else 0.58))
            y = round(image.height * 0.30)
            draw.arc((x - 5, y - 5, x + 8, y + 6), start=70 if side < 0 else 200, end=160 if side < 0 else 290, fill=(80, 70, 60, 150), width=1)

    def _transform_subject(self, image: Image.Image, dx: int, dy: int, angle: float) -> Image.Image:
        bbox = image.getbbox()
        if not bbox:
            return image
        subject = image.crop(bbox)
        subject = subject.rotate(angle, resample=Image.Resampling.BICUBIC, expand=True)
        canvas = Image.new("RGBA", image.size, (0, 0, 0, 0))
        x = round((bbox[0] + bbox[2] - subject.width) / 2 + dx)
        y = round((bbox[1] + bbox[3] - subject.height) / 2 + dy)
        canvas.alpha_composite(subject, (x, y))
        return canvas

    def _draw_lab_focus(self, image: Image.Image, index: int) -> None:
        draw = ImageDraw.Draw(image, "RGBA")
        if index in {1, 2, 3, 7, 8, 9}:
            x = round(image.width * 0.18)
            y = round(image.height * 0.47)
            draw.line((x, y, x - 8, y + 20), fill=(65, 86, 138, 180), width=2)
        if index in {3, 4, 5}:
            cx = round(image.width * 0.35)
            cy = round(image.height * 0.36)
            draw.arc((cx - 10, cy - 5, cx + 8, cy + 10), start=210, end=320, fill=(255, 210, 90, 170), width=2)

    def _draw_typing_hands(self, image: Image.Image, index: int) -> None:
        draw = ImageDraw.Draw(image, "RGBA")
        points = [
            (0.38, 0.69, index % 4 in {0, 1}),
            (0.50, 0.68, index % 4 in {2, 3}),
        ]
        for rx, ry, down in points:
            x = round(image.width * rx)
            y = round(image.height * ry) + (3 if down else -1)
            draw.ellipse((x - 3, y - 2, x + 4, y + 3), fill=(246, 190, 154, 240), outline=(74, 50, 44, 160))

    def _draw_closed_eye_smile(self, image: Image.Image) -> None:
        draw = ImageDraw.Draw(image, "RGBA")
        skin = (250, 197, 160, 235)
        dark = (45, 35, 30, 230)
        left_eye = self._box(image, (0.35, 0.42, 0.43, 0.48))
        right_eye = self._box(image, (0.47, 0.42, 0.55, 0.48))
        for box in (left_eye, right_eye):
            draw.rounded_rectangle(box, radius=2, fill=skin)
            draw.arc((box[0] - 1, box[1] - 3, box[2] + 1, box[3] + 5), start=20, end=160, fill=dark, width=2)
        mouth = self._box(image, (0.43, 0.51, 0.49, 0.55))
        draw.arc((mouth[0], mouth[1] - 2, mouth[2], mouth[3] + 2), start=20, end=160, fill=(190, 80, 72, 230), width=1)

    def _draw_phone_taps(self, image: Image.Image, index: int) -> None:
        draw = ImageDraw.Draw(image, "RGBA")
        x = round(image.width * 0.43)
        y = round(image.height * 0.57)
        if index % 3 == 0:
            draw.ellipse((x - 2, y - 2, x + 2, y + 2), fill=(210, 235, 255, 150))
        if index % 4 in {1, 2}:
            draw.line((x - 5, y + 7, x + 8, y + 7), fill=(210, 235, 255, 120), width=1)

    def _typing_ticks(self, image: Image.Image, index: int) -> None:
        draw = ImageDraw.Draw(image, "RGBA")
        base_x = round(image.width * 0.36)
        base_y = round(image.height * 0.68)
        for dot in range(3):
            if (index + dot) % 3 == 0:
                x = base_x + dot * 4
                draw.rectangle((x, base_y, x + 2, base_y + 1), fill=(245, 250, 255, 180))

    def _liquid_pulse(self, image: Image.Image, index: int) -> None:
        draw = ImageDraw.Draw(image, "RGBA")
        phase = index % 6
        if phase in {1, 2, 3}:
            x = round(image.width * 0.18)
            y = round(image.height * (0.60 + phase * 0.025))
            draw.ellipse((x - 1, y - 2, x + 2, y + 2), fill=(70, 210, 220, 180))

    def _beaker_bubbles(self, image: Image.Image, index: int) -> None:
        draw = ImageDraw.Draw(image, "RGBA")
        offsets = ((0.13, 0.76), (0.18, 0.72), (0.24, 0.78))
        for bubble_index, (rx, ry) in enumerate(offsets):
            if (index + bubble_index) % 4 in {0, 1}:
                x = round(image.width * rx)
                y = round(image.height * ry - (index % 4) * 2)
                draw.ellipse((x - 1, y - 1, x + 2, y + 2), outline=(120, 230, 235, 170))

    def _is_mostly_static(self, frames: list[Image.Image]) -> bool:
        if len(frames) < 2:
            return True

        differences = []
        for current, next_frame in zip(frames, frames[1:]):
            diff = ImageChops.difference(current, next_frame)
            differences.append(sum(ImageStat.Stat(diff).sum))
        return max(differences, default=0) < 100_000

    def _synthesize_motion(
        self,
        base: Image.Image,
        count: int,
        bob: float,
        squash: float,
        tilt: float,
    ) -> list[Image.Image]:
        result = []
        width, height = base.size
        for index in range(count):
            phase = math.sin(index / count * math.tau)
            secondary = math.sin(index / count * math.tau + math.pi / 2)
            scale_x = 1 + squash * max(0, -phase)
            scale_y = 1 - squash * max(0, -phase)
            resized = base.resize(
                (max(1, round(width * scale_x)), max(1, round(height * scale_y))),
                Image.Resampling.LANCZOS,
            )
            rotated = resized.rotate(tilt * secondary, resample=Image.Resampling.BICUBIC, expand=True)
            canvas = Image.new("RGBA", base.size, (0, 0, 0, 0))
            x = (width - rotated.width) // 2
            y = round((height - rotated.height) / 2 + bob * phase)
            canvas.alpha_composite(rotated, (x, y))
            result.append(self._remove_halo(canvas))
        return result

    def _build_menu(self) -> Menu:
        menu = Menu(self.root, tearoff=0)
        menu.add_command(label="待机", command=lambda: self._set_free_state("idle"))
        menu.add_command(label="思考", command=lambda: self._set_free_state("think"))
        menu.add_separator()

        for action_key, label in TIMED_ACTIONS.items():
            submenu = Menu(menu, tearoff=0)
            for text, seconds in TIMES.items():
                submenu.add_command(
                    label=text,
                    command=lambda key=action_key, duration=seconds: self._set_timed_state(key, duration),
                )
            menu.add_cascade(label=label, menu=submenu)

        menu.add_separator()
        menu.add_command(label="挥手", command=lambda: self._play_once("wave"))
        menu.add_separator()
        size_menu = Menu(menu, tearoff=0)
        for label, scale in SIZE_OPTIONS.items():
            size_menu.add_command(label=label, command=lambda value=scale: self._set_scale(value))
        menu.add_cascade(label="大小", menu=size_menu)
        menu.add_separator()
        menu.add_command(label="退出", command=self.root.destroy)
        return menu

    def _place_initially(self) -> None:
        self.root.update_idletasks()
        width = self.root.winfo_reqwidth()
        height = self.root.winfo_reqheight()
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = screen_width - width - 80
        y = screen_height - height - 90
        self.root.geometry(f"+{max(0, x)}+{max(0, y)}")

    def _set_free_state(self, action_key: str) -> None:
        self._cancel_timer()
        state = ACTION_STATES[action_key]
        self.return_state = state
        self._set_state(state)

    def _set_timed_state(self, action_key: str, seconds: int) -> None:
        self._cancel_timer()
        self.one_shot_state = None
        state = ACTION_STATES[action_key]
        self.return_state = state
        self._set_state(state)
        self.timer_job = self.root.after(seconds * 1000, self._finish_timed_action)

    def _finish_timed_action(self) -> None:
        self.timer_job = None
        self.one_shot_state = None
        self.return_state = "idle"
        if not self.dragging:
            self._set_state("idle")

    def _play_once(self, action_key: str) -> None:
        self._cancel_timer()
        state = ACTION_STATES[action_key]
        self.one_shot_state = state
        self.return_state = "idle"
        self._set_state(state)
        frame_count = len(self.frames.get(state, []))
        self.timer_job = self.root.after(max(900, frame_count * 130), self._finish_timed_action)

    def _cancel_timer(self) -> None:
        if self.timer_job is not None:
            self.root.after_cancel(self.timer_job)
            self.timer_job = None
        self.one_shot_state = None

    def _set_scale(self, scale: float) -> None:
        if scale == self.scale:
            return
        self.scale = scale
        self.frames = self._load_frames(scale)
        self.frame_index = 0
        state_frames = self.frames.get(self.state) or self.frames["idle"]
        self.label.configure(image=state_frames[0])

    def _set_state(self, state: str) -> None:
        if state not in self.frames:
            state = "idle"
        if self.state != state:
            self.state = state
            self.frame_index = 0

    def _tick(self) -> None:
        state_frames = self.frames.get(self.state) or self.frames["idle"]
        self.label.configure(image=state_frames[self.frame_index % len(state_frames)])
        self.frame_index += 1
        self.root.after(120, self._tick)

    def _start_drag(self, event: tk.Event) -> None:
        self.pointer_down = True
        self.dragging = False
        self.press_button = event.num
        self.drag_offset_x = event.x
        self.drag_offset_y = event.y
        self.press_root_x = event.x_root
        self.press_root_y = event.y_root
        self.last_pointer_x = event.x_root
        self.drag_state = "running-right"

    def _right_press(self, event: tk.Event) -> None:
        self._start_drag(event)

    def _drag(self, event: tk.Event) -> None:
        if not self.pointer_down:
            return

        distance = abs(event.x_root - self.press_root_x) + abs(event.y_root - self.press_root_y)
        if not self.dragging and distance < 6:
            return

        if self.click_job is not None:
            self.root.after_cancel(self.click_job)
            self.click_job = None
        if self.one_shot_state is not None:
            self._cancel_timer()
            self.return_state = "idle"

        self.dragging = True
        x = event.x_root - self.drag_offset_x
        y = event.y_root - self.drag_offset_y
        self.root.geometry(f"+{x}+{y}")

        delta_x = event.x_root - self.last_pointer_x
        self.last_pointer_x = event.x_root
        if delta_x < -1:
            self.drag_state = "running-left"
        elif delta_x > 1:
            self.drag_state = "running-right"
        self._set_state(self.drag_state)

    def _stop_drag(self, event: tk.Event) -> None:
        was_dragging = self.dragging
        self.pointer_down = False
        self.dragging = False
        if was_dragging:
            self._set_state(self.return_state)
        else:
            if self._is_duplicate_left_release(event):
                return
            self.last_left_release_time = event.time
            self.last_left_release_x = event.x_root
            self.last_left_release_y = event.y_root
            self._handle_left_click()

    def _right_release(self, event: tk.Event) -> None:
        was_dragging = self.dragging
        moved = abs(event.x_root - self.press_root_x) + abs(event.y_root - self.press_root_y)
        self.pointer_down = False
        self.dragging = False
        if was_dragging:
            self._set_state(self.return_state)
        elif moved < 6:
            self.menu.tk_popup(event.x_root, event.y_root)
            self.menu.grab_release()

    def _handle_left_click(self) -> None:
        if self.click_job is not None:
            self.root.after_cancel(self.click_job)
            self.click_job = None
            self._handle_double_click()
            return

        self.click_job = self.root.after(360, self._handle_single_click)

    def _is_duplicate_left_release(self, event: tk.Event) -> bool:
        if event.time - self.last_left_release_time > 80:
            return False
        moved = abs(event.x_root - self.last_left_release_x) + abs(event.y_root - self.last_left_release_y)
        return moved < 4

    def _handle_single_click(self) -> None:
        self.click_job = None
        self._play_once("think")

    def _handle_double_click(self) -> None:
        self._play_once("jump")
        webbrowser.open(ACTA_URL, new=2)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Codex 桌面宠物")
    parser.add_argument("--scale", type=float, default=0.8, help="宠物缩放比例，默认 0.8。")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    if args.scale <= 0:
        print("--scale 必须大于 0。", file=sys.stderr)
        return 2

    try:
        CodexDesktopPet(args.scale).run()
    except Exception as exc:
        print(f"启动失败：{exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
