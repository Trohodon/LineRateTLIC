from __future__ import annotations

import random
import tkinter as tk
import time


class SessionPanel(tk.Toplevel):
    _best_score = 0

    def __init__(self, parent: tk.Misc) -> None:
        super().__init__(parent)
        self.title("Session Panel")
        self.resizable(False, False)
        self.transient(parent)

        self._cell = 18
        self._cols = 30
        self._rows = 22
        self._base_speed_ms = 86
        self._min_speed_ms = 26
        self._speed_step_ms = 4

        self._heading = (1, 0)
        self._next_heading = (1, 0)
        self._turn_queue: list[tuple[int, int]] = []
        self._trail = [(10, 11), (9, 11), (8, 11)]
        self._target = (15, 11)
        self._bonus_target: tuple[int, int] | None = None
        self._bonus_ticks = 0
        self._running = False
        self._state = "ready"
        self._score = 0
        self._next_tick_at = 0.0
        self._frame = 0
        self._stars = [
            (
                random.randint(0, self._cols * self._cell - 1),
                random.randint(0, self._rows * self._cell - 1),
                random.randint(1, 3),
                random.randint(0, 7),
            )
            for _ in range(90)
        ]
        self._after_id: str | None = None

        width = self._cols * self._cell
        height = self._rows * self._cell

        self._canvas = tk.Canvas(self, width=width, height=height, background="#0f1418", highlightthickness=0)
        self._canvas.grid(row=0, column=0, padx=10, pady=(10, 6))

        self._line = tk.StringVar(value="")
        self._meta = tk.StringVar(value="")
        label = tk.Label(self, textvariable=self._line, background="#f0f0f0", foreground="#343a40", anchor="w")
        label.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 2))
        meta = tk.Label(self, textvariable=self._meta, background="#f0f0f0", foreground="#747f8a", anchor="w")
        meta.grid(row=2, column=0, sticky="ew", padx=10, pady=(0, 10))

        self.bind_all("<KeyPress>", self._on_keypress, add="+")
        self.protocol("WM_DELETE_WINDOW", self._close)

        self._seed_target()
        self._refresh_labels()
        self._draw_static()
        self._draw()
        self.focus_force()

    def _on_keypress(self, event: tk.Event) -> None:
        focused = self.focus_get()
        if focused is None or focused.winfo_toplevel() is not self:
            return
        key = (event.keysym or "").lower()
        if key in {"up", "w"}:
            self._set_heading(0, -1)
        elif key in {"down", "s"}:
            self._set_heading(0, 1)
        elif key in {"left", "a"}:
            self._set_heading(-1, 0)
        elif key in {"right", "d"}:
            self._set_heading(1, 0)
        elif key in {"return", "space", "p"}:
            self._toggle()

    def _toggle(self) -> None:
        if self._state in {"ready", "game_over"}:
            self._start_new_run()
            return
        if self._running and self._state == "running":
            self._running = False
            self._state = "paused"
            if self._after_id is not None:
                self.after_cancel(self._after_id)
                self._after_id = None
            self._refresh_labels()
            return
        if self._state == "paused":
            self._running = True
            self._state = "running"
            self._next_tick_at = time.perf_counter()
            self._refresh_labels()
            self._tick()

    def _set_heading(self, dx: int, dy: int) -> None:
        if dx == 0 and dy == 0:
            return
        # Compare against the latest queued direction so rapid key taps still
        # form clean turns and never allow an immediate reversal.
        anchor = self._turn_queue[-1] if self._turn_queue else self._next_heading
        if (dx, dy) == (-anchor[0], -anchor[1]):
            return
        if self._turn_queue and self._turn_queue[-1] == (dx, dy):
            return
        if len(self._turn_queue) < 3:
            self._turn_queue.append((dx, dy))
        else:
            self._turn_queue[-1] = (dx, dy)

    def _start_new_run(self) -> None:
        self._reset_state()
        self._running = True
        self._state = "running"
        self._next_tick_at = time.perf_counter()
        self._refresh_labels()
        self._draw()
        self._tick()

    def _tick(self) -> None:
        if not self._running or self._state != "running":
            return
        if self._turn_queue:
            self._next_heading = self._turn_queue.pop(0)
        self._heading = self._next_heading
        head_x, head_y = self._trail[0]
        nx = head_x + self._heading[0]
        ny = head_y + self._heading[1]

        if nx < 0 or ny < 0 or nx >= self._cols or ny >= self._rows or (nx, ny) in self._trail:
            self._running = False
            self._state = "game_over"
            if self._score > SessionPanel._best_score:
                SessionPanel._best_score = self._score
            self._refresh_labels()
            self._draw()
            return

        self._trail.insert(0, (nx, ny))
        grew = False
        if (nx, ny) == self._target:
            grew = True
            self._score += 1
            self._seed_target()
            if self._score % 5 == 0 and self._bonus_target is None:
                self._seed_bonus()
        elif self._bonus_target is not None and (nx, ny) == self._bonus_target:
            grew = True
            self._score += 3
            self._bonus_target = None
            self._bonus_ticks = 0
        if not grew:
            self._trail.pop()

        if self._bonus_target is not None:
            self._bonus_ticks -= 1
            if self._bonus_ticks <= 0:
                self._bonus_target = None

        if self._score > SessionPanel._best_score:
            SessionPanel._best_score = self._score

        self._refresh_labels()
        self._draw()
        speed_ms = max(self._min_speed_ms, self._base_speed_ms - self._score * self._speed_step_ms)
        interval = speed_ms / 1000.0
        now = time.perf_counter()
        if self._next_tick_at <= 0.0:
            self._next_tick_at = now + interval
        else:
            self._next_tick_at += interval
            if self._next_tick_at < now:
                self._next_tick_at = now + interval
        delay = max(1, int((self._next_tick_at - now) * 1000))
        self._after_id = self.after(delay, self._tick)

    def _reset_state(self) -> None:
        self._heading = (1, 0)
        self._next_heading = (1, 0)
        self._turn_queue.clear()
        self._trail = [(10, 11), (9, 11), (8, 11)]
        self._score = 0
        self._bonus_target = None
        self._bonus_ticks = 0
        self._seed_target()

    def _seed_target(self) -> None:
        blocked = set(self._trail)
        if self._bonus_target is not None:
            blocked.add(self._bonus_target)
        open_cells = [(x, y) for x in range(self._cols) for y in range(self._rows) if (x, y) not in blocked]
        if not open_cells:
            self._target = (0, 0)
            return
        self._target = random.choice(open_cells)

    def _seed_bonus(self) -> None:
        blocked = set(self._trail)
        blocked.add(self._target)
        open_cells = [(x, y) for x in range(self._cols) for y in range(self._rows) if (x, y) not in blocked]
        if not open_cells:
            return
        self._bonus_target = random.choice(open_cells)
        self._bonus_ticks = 70

    def _refresh_labels(self) -> None:
        if self._state == "ready":
            self._line.set("Press Enter to start")
            self._meta.set(f"Best: {SessionPanel._best_score}")
            return
        if self._state == "paused":
            self._line.set(f"Paused  |  Score: {self._score}")
            self._meta.set(f"Best: {SessionPanel._best_score}  |  Enter/P to resume")
            return
        if self._state == "game_over":
            self._line.set(f"Game Over  |  Score: {self._score}")
            self._meta.set(f"Best: {SessionPanel._best_score}  |  Enter to restart")
            return
        self._line.set(f"Score: {self._score}")
        self._meta.set(f"Best: {SessionPanel._best_score}  |  Enter/P to pause")

    def _draw_static(self) -> None:
        c = self._canvas
        c.delete("static")
        w = self._cols * self._cell
        h = self._rows * self._cell
        c.create_rectangle(0, 0, w, h, fill="#0e1419", outline="", tags="static")
        c.create_rectangle(1, 1, w - 1, h - 1, outline="#2a3239", tags="static")
        for x in range(0, self._cols, 2):
            x0 = x * self._cell
            c.create_line(x0, 0, x0, h, fill="#151c22", tags="static")
        for y in range(0, self._rows, 2):
            y0 = y * self._cell
            c.create_line(0, y0, w, y0, fill="#151c22", tags="static")

    def _draw(self) -> None:
        c = self._canvas
        c.delete("dynamic")
        self._frame += 1
        w = self._cols * self._cell
        h = self._rows * self._cell

        # Moving background bands for motion feel.
        phase = (self._frame * 3) % 48
        for i in range(-1, (h // 24) + 3):
            y0 = i * 24 + phase - 24
            c.create_rectangle(0, y0, w, y0 + 10, fill="#111a21", outline="", tags="dynamic")

        # Subtle starfield twinkle.
        for sx, sy, r, tw in self._stars:
            if ((self._frame + tw) % 18) < 8:
                c.create_oval(sx - r, sy - r, sx + r, sy + r, fill="#1f2b34", outline="", tags="dynamic")
            else:
                c.create_oval(sx, sy, sx + 1, sy + 1, fill="#31424f", outline="", tags="dynamic")

        tx, ty = self._target
        c.create_oval(
            tx * self._cell + 2,
            ty * self._cell + 2,
            (tx + 1) * self._cell - 2,
            (ty + 1) * self._cell - 2,
            fill="#cfd8dc",
            outline="#90a4ae",
            tags="dynamic",
        )

        if self._bonus_target is not None:
            bx, by = self._bonus_target
            pad = 3
            x0 = bx * self._cell + pad
            y0 = by * self._cell + pad
            x1 = (bx + 1) * self._cell - pad
            y1 = (by + 1) * self._cell - pad
            mx = (x0 + x1) / 2
            my = (y0 + y1) / 2
            c.create_polygon(mx, y0, x1, my, mx, y1, x0, my, fill="#ffd166", outline="#e09f3e", tags="dynamic")

        for i, (x, y) in enumerate(self._trail):
            color = "#82d173" if i == 0 else "#3f8f50"
            c.create_rectangle(
                x * self._cell + 2,
                y * self._cell + 2,
                (x + 1) * self._cell - 2,
                (y + 1) * self._cell - 2,
                fill=color,
                outline="#2f5f35",
                tags="dynamic",
            )
        if self._trail:
            hx, hy = self._trail[0]
            cx = hx * self._cell + self._cell / 2
            cy = hy * self._cell + self._cell / 2
            ox = self._heading[0] * 3
            oy = self._heading[1] * 3
            c.create_oval(cx - 4 + ox, cy - 4 + oy, cx - 1 + ox, cy - 1 + oy, fill="#0f1418", outline="", tags="dynamic")
            c.create_oval(cx + 1 + ox, cy - 4 + oy, cx + 4 + ox, cy - 1 + oy, fill="#0f1418", outline="", tags="dynamic")

    def _close(self) -> None:
        self._running = False
        if self._after_id is not None:
            self.after_cancel(self._after_id)
            self._after_id = None
        self.unbind_all("<KeyPress>")
        self.destroy()


def open_panel(parent: tk.Misc) -> None:
    SessionPanel(parent)
