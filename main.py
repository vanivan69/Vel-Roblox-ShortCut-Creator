import os
import sys
import uuid
import asyncio
import io
import re
from ctypes import windll, wintypes
import aiohttp
from PIL import Image, ImageDraw, ImageFont
import customtkinter as ctk
import win32com.client
import pythoncom

pythoncom.CoInitialize()

ctk.set_appearance_mode("Dark")

# Win32 API Constants
GWL_EXSTYLE = -20
WS_EX_APPWINDOW = 0x00040000
WS_EX_TOOLWINDOW = 0x00000080
FR_PRIVATE = 0x10

# Color Palette
COLOR_BG = "#0B0F19"
COLOR_SURFACE = "#151C2C"
COLOR_SURFACE_HOVER = "#1E293B"
COLOR_BORDER = "#2A364F"
COLOR_TEXT_MAIN = "#F1F5F9"
COLOR_TEXT_MUTED = "#64748B"
COLOR_ACCENT = "#38BDF8"
COLOR_ERROR = "#F87171"
COLOR_SUCCESS = "#4ADE80"
COLOR_TRANSPARENT_KEY = "#000001"

REQUEST_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "en-US,en;q=0.9",
    "Referer": "https://www.roblox.com/"
}


def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    
    exe_dir = os.path.dirname(os.path.abspath(sys.executable))
    path_near_exe = os.path.join(exe_dir, relative_path)
    if os.path.exists(path_near_exe):
        return path_near_exe
        
    return os.path.join(os.path.abspath("."), relative_path)


def load_custom_font():
    font_path = resource_path("Metropolis-ExtraBold.otf")
    if os.path.exists(font_path):
        res = windll.gdi32.AddFontResourceExW(font_path, FR_PRIVATE, 0)
        if res > 0:
            return "Metropolis Extra Bold"
    return "Segoe UI"


FONT_FAMILY = load_custom_font()


def round_image_corners(pil_img, radius=14):
    pil_img = pil_img.convert("RGBA")
    mask = Image.new("L", pil_img.size, 0)
    draw = ImageDraw.Draw(mask)
    draw.rounded_rectangle((0, 0, pil_img.size[0], pil_img.size[1]), radius=radius, fill=255)
    pil_img.putalpha(mask)
    return pil_img


def render_text_with_color_emojis(text, font_size=13, text_color=(241, 245, 249, 255), max_width=180):
    try:
        emoji_font = ImageFont.truetype("seguiemj.ttf", int(font_size * 0.95))
    except IOError:
        emoji_font = ImageFont.load_default()

    try:
        font_file = resource_path("Metropolis-ExtraBold.otf")
        main_font = ImageFont.truetype(font_file, font_size) if os.path.exists(font_file) else ImageFont.load_default()
    except IOError:
        main_font = ImageFont.load_default()

    emoji_pattern = re.compile(
        r'([\U00010000-\U0010ffff\u2600-\u27ff\u2300-\u23ff\u2B05\u2194-\u2199\u2B06\u2B07])'
    )
    tokens = [t for t in emoji_pattern.split(text) if t]

    dummy_img = Image.new("RGBA", (1, 1))
    dummy_draw = ImageDraw.Draw(dummy_img)

    lines = []
    current_line = []
    current_w = 0

    for token in tokens:
        is_emoji = bool(emoji_pattern.match(token))
        f = emoji_font if is_emoji else main_font
        bbox = dummy_draw.textbbox((0, 0), token, font=f, embedded_color=is_emoji)
        token_w = bbox[2] - bbox[0]

        if current_w + token_w > max_width and current_line:
            lines.append(current_line)
            current_line = [(token, is_emoji, f, token_w)]
            current_w = token_w
        else:
            current_line.append((token, is_emoji, f, token_w))
            current_w += token_w

    if current_line:
        lines.append(current_line)

    lines = lines[:2]
    line_height = int(font_size * 1.4)
    canvas_h = max(30, len(lines) * line_height)
    canvas_w = max_width

    img = Image.new("RGBA", (canvas_w, canvas_h), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    y = 0
    for line in lines:
        x = 0
        for token, is_emoji, f, token_w in line:
            if is_emoji:
                draw.text((x, y + 2), token, font=f, embedded_color=True)
            else:
                draw.text((x, y), token, font=f, fill=text_color)
            x += token_w
        y += line_height

    return ctk.CTkImage(light_image=img, dark_image=img, size=(canvas_w, canvas_h))


class RobloxShortcutApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.APP_NAME = "Vel's Roblox SCC"
        self.title(self.APP_NAME)
        self.overrideredirect(True)
        self.geometry("780x580")

        self.configure(fg_color=COLOR_TRANSPARENT_KEY)
        self.wm_attributes("-transparentcolor", COLOR_TRANSPARENT_KEY)

        self.session_id = str(uuid.uuid4())
        self.suggestion_text = ""
        self.debounce_timer = None
        self.cards_map = {}
        self._drag_data = {"x": 0, "y": 0}

        self.loop = asyncio.new_event_loop()
        asyncio.set_event_loop(self.loop)

        self._build_ui()
        self._bind_events()

        self.after(10, self._set_app_identity)
        self.after(20, self._process_asyncio)

    def _get_text_width(self, text, font_size=13):
        try:
            font_file = resource_path("Metropolis-ExtraBold.otf")
            font = ImageFont.truetype(font_file, font_size) if os.path.exists(font_file) else ImageFont.load_default()
            return int(font.getlength(text))
        except Exception:
            return len(text) * 8

    def log_status(self, message, color=COLOR_TEXT_MUTED):
        print(f"[LOG] {message}")
        self.after(0, lambda: self.status_label.configure(text=message, text_color=color))

    def _set_app_identity(self):
        try:
            ico_path = resource_path("app.ico")
            if os.path.exists(ico_path):
                self.iconbitmap(ico_path)

            hwnd = windll.user32.GetParent(self.winfo_id())
            if not hwnd:
                hwnd = self.winfo_id()

            windll.user32.SetWindowTextW(hwnd, self.APP_NAME)

            style = windll.user32.GetWindowLongPtrW(hwnd, GWL_EXSTYLE)
            style = style & ~WS_EX_TOOLWINDOW
            style = style | WS_EX_APPWINDOW
            windll.user32.SetWindowLongPtrW(hwnd, GWL_EXSTYLE, style)

            self.withdraw()
            self.deiconify()
        except Exception as e:
            print(f"Taskbar identity error: {e}")

    def _build_ui(self):
        self.main_container = ctk.CTkFrame(
            self,
            fg_color=COLOR_BG,
            border_color=COLOR_BORDER,
            border_width=1,
            corner_radius=18
        )
        self.main_container.pack(fill="both", expand=True, padx=2, pady=2)

        self.title_bar = ctk.CTkFrame(
            self.main_container,
            fg_color="transparent",
            height=42,
            corner_radius=0
        )
        self.title_bar.pack(fill="x", padx=14, pady=(10, 0))

        ico_path = resource_path("app.ico")
        if os.path.exists(ico_path):
            ico_pil = Image.open(ico_path).resize((22, 22))
            self.app_icon_ctk = ctk.CTkImage(light_image=ico_pil, dark_image=ico_pil, size=(22, 22))
            self.app_icon_label = ctk.CTkLabel(self.title_bar, image=self.app_icon_ctk, text="")
            self.app_icon_label.pack(side="left", padx=(4, 8))

        self.title_label = ctk.CTkLabel(
            self.title_bar,
            text=self.APP_NAME,
            font=(FONT_FAMILY, 15, "bold"),
            text_color=COLOR_TEXT_MAIN
        )
        self.title_label.pack(side="left")

        self.close_btn = ctk.CTkButton(
            self.title_bar, text="✕", width=28, height=28, corner_radius=14,
            fg_color="transparent", hover_color="#EF4444", text_color=COLOR_TEXT_MAIN,
            font=(FONT_FAMILY, 12, "bold"), command=self.destroy
        )
        self.close_btn.pack(side="right", padx=2)

        self.max_btn = ctk.CTkButton(
            self.title_bar, text="☐", width=28, height=28, corner_radius=14,
            fg_color="transparent", hover_color=COLOR_SURFACE, text_color=COLOR_TEXT_MAIN,
            font=(FONT_FAMILY, 12, "bold")
        )
        self.max_btn.pack(side="right", padx=2)

        self.min_btn = ctk.CTkButton(
            self.title_bar, text="—", width=28, height=28, corner_radius=14,
            fg_color="transparent", hover_color=COLOR_SURFACE, text_color=COLOR_TEXT_MAIN,
            font=(FONT_FAMILY, 12, "bold"), command=self._minimize_window
        )
        self.min_btn.pack(side="right", padx=2)

        self.title_bar.bind("<ButtonPress-1>", self._start_drag)
        self.title_bar.bind("<B1-Motion>", self._do_drag)
        self.title_label.bind("<ButtonPress-1>", self._start_drag)
        self.title_label.bind("<B1-Motion>", self._do_drag)

        self.search_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        self.search_frame.pack(fill="x", padx=20, pady=(15, 10))

        self.search_box_container = ctk.CTkFrame(
            self.search_frame,
            fg_color=COLOR_SURFACE,
            border_color=COLOR_BORDER,
            border_width=1,
            corner_radius=12,
            height=42
        )
        self.search_box_container.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.search_box_container.pack_propagate(False)

        self.search_icon = ctk.CTkLabel(
            self.search_box_container,
            text="🔍",
            font=("Segoe UI", 13),
            text_color=COLOR_TEXT_MUTED
        )
        self.search_icon.pack(side="left", padx=(12, 4))

        self.input_wrapper = ctk.CTkFrame(self.search_box_container, fg_color="transparent")
        self.input_wrapper.pack(side="left", fill="both", expand=True, padx=(0, 10))

        self.suggestion_label = ctk.CTkLabel(
            self.input_wrapper,
            text="",
            font=(FONT_FAMILY, 13),
            text_color=COLOR_TEXT_MUTED,
            anchor="w"
        )
        self.suggestion_label.bind("<Button-1>", lambda e: self.search_entry.focus_set())

        self.search_entry = ctk.CTkEntry(
            self.input_wrapper,
            placeholder_text="Search game or enter Place ID...",
            placeholder_text_color=COLOR_TEXT_MUTED,
            font=(FONT_FAMILY, 13),
            fg_color="transparent",
            border_width=0,
            text_color=COLOR_TEXT_MAIN
        )
        self.search_entry.place(relx=0, rely=0, relwidth=1, relheight=1)

        self.search_button = ctk.CTkButton(
            self.search_frame,
            text="Search",
            width=110,
            height=42,
            corner_radius=12,
            fg_color=COLOR_SURFACE,
            hover_color=COLOR_SURFACE_HOVER,
            border_color=COLOR_BORDER,
            border_width=1,
            font=(FONT_FAMILY, 13, "bold"),
            text_color=COLOR_TEXT_MAIN,
            command=self.on_search_click
        )
        self.search_button.pack(side="right")

        self.results_outer_frame = ctk.CTkFrame(
            self.main_container,
            fg_color=COLOR_SURFACE,
            border_color=COLOR_BORDER,
            border_width=1,
            corner_radius=14
        )
        self.results_outer_frame.pack(fill="both", expand=True, padx=20, pady=5)

        self.results_scrollable = ctk.CTkScrollableFrame(
            self.results_outer_frame,
            orientation="vertical",
            fg_color="transparent",
            scrollbar_button_color=COLOR_BORDER,
            scrollbar_button_hover_color=COLOR_ACCENT
        )
        self.results_scrollable.pack(fill="both", expand=True, padx=8, pady=8)
        self.results_scrollable.grid_columnconfigure((0, 1), weight=1)

        self.error_frame = ctk.CTkFrame(self.main_container, fg_color="transparent", height=24)
        self.error_frame.pack(fill="x", padx=20, pady=(4, 10))

        self.status_label = ctk.CTkLabel(
            self.error_frame,
            text="Ready.",
            font=(FONT_FAMILY, 12),
            text_color=COLOR_TEXT_MUTED
        )
        self.status_label.pack(side="bottom")

    def _bind_events(self):
        self.search_entry.bind("<Tab>", self.on_tab_press)
        self.search_entry.bind("<KeyRelease>", self.on_key_release)
        self.search_entry.bind("<Return>", lambda e: self.on_search_click())

    def _minimize_window(self):
        self.update_idletasks()
        self.overrideredirect(False)
        self.iconify()
        self.after(100, lambda: self.bind("<FocusIn>", self._restore_override))

    def _restore_override(self, event=None):
        self.unbind("<FocusIn>")
        self.overrideredirect(True)
        self._set_app_identity()

    def _start_drag(self, event):
        self._drag_data["x"] = event.x
        self._drag_data["y"] = event.y

    def _do_drag(self, event):
        x = self.winfo_x() + (event.x - self._drag_data["x"])
        y = self.winfo_y() + (event.y - self._drag_data["y"])
        self.geometry(f"+{x}+{y}")

    def _process_asyncio(self):
        self.loop.stop()
        self.loop.run_forever()
        self.after(20, self._process_asyncio)

    def async_task(self, coro):
        def _start():
            asyncio.create_task(coro)
        self.loop.call_soon_threadsafe(_start)

    def get_session(self):
        return aiohttp.ClientSession(headers=REQUEST_HEADERS)

    def on_tab_press(self, event):
        if self.suggestion_text:
            self.search_entry.delete(0, "end")
            self.search_entry.insert(0, self.suggestion_text)
            self.search_entry.icursor("end")
            self.suggestion_text = ""
            self.suggestion_label.configure(text="")
            self.log_status(f"Accepted suggestion: '{self.search_entry.get()}'", COLOR_SUCCESS)
            return "break"

    def on_key_release(self, event):
        if event.keysym in ("Tab", "Return", "Up", "Down", "Left", "Right", "Shift_L", "Shift_R", "Control_L", "Control_R"):
            return

        self.suggestion_text = ""
        self.suggestion_label.configure(text="")

        if self.debounce_timer:
            self.after_cancel(self.debounce_timer)
            self.debounce_timer = None

        raw_input = self.search_entry.get()
        clean_text = raw_input.strip()

        if not clean_text or clean_text.isdigit():
            self.log_status("Ready.", COLOR_TEXT_MUTED)
            return

        self.debounce_timer = self.after(300, lambda: self.async_task(self.fetch_autocomplete(clean_text)))

    async def fetch_autocomplete(self, search_text):
        query_text = search_text.lower()
        url = f"https://apis.roblox.com/games-autocomplete/v1/get-suggestion/{query_text}"
        
        try:
            async with self.get_session() as session:
                async with session.get(url, timeout=3) as response:
                    if response.status == 200:
                        data = await response.json()
                        entries = data.get("entries", [])
                        if entries:
                            first_entry = entries[0]
                            best_match = first_entry.get("searchQuery", "")
                            best_match_lower = best_match.lower()

                            if best_match_lower == query_text:
                                self.suggestion_text = ""
                                self.after(0, self._update_suggestion_ui, search_text, "")
                                self.log_status("Ready.", COLOR_TEXT_MUTED)
                                return

                            if best_match_lower.startswith(query_text) and len(best_match) > len(search_text):
                                self.suggestion_text = best_match
                                self.after(0, self._update_suggestion_ui, search_text, best_match)
                                self.log_status(f"Maybe you've meant {best_match}?", COLOR_ACCENT)
                                return
                    else:
                        self.log_status(f"Autocomplete API status: {response.status}", COLOR_ERROR)
        except Exception as e:
            self.log_status(f"Autocomplete Error: {e}", COLOR_ERROR)

        self.suggestion_text = ""
        self.after(0, self._update_suggestion_ui, "", "")   

    def _update_suggestion_ui(self, search_text, best_match):
        current_input = self.search_entry.get()
        
        if (
            best_match 
            and current_input 
            and best_match.lower().startswith(current_input.lower())
            and len(best_match) > len(current_input)
        ):
            remainder = best_match[len(current_input):]
            text_offset = self._get_text_width(current_input, font_size=13) + 5
            
            self.suggestion_label.place(x=text_offset, y=-1.1, relheight=1)
            self.suggestion_label.configure(text=remainder, text_color=COLOR_TEXT_MUTED)
            self.suggestion_label.lift()
        else:
            self.suggestion_label.configure(text="")

    def on_search_click(self):
        raw_query = self.search_entry.get().strip()
        if not raw_query:
            self.log_status("Please enter a game name or Place ID.", COLOR_ERROR)
            return

        self.search_button.configure(state="disabled")
        self.async_task(self.perform_search(raw_query))

    async def perform_search(self, query):
        results = []
        
        place_id_match = re.search(r'\b\d{5,20}\b', query)

        if place_id_match:
            place_id = int(place_id_match.group(0))
            self.log_status(f"Detected Place ID: {place_id}. Fetching details...", COLOR_ACCENT)
            
            # Запрос данных плейса по его ID
            details_url = f"https://games.roblox.com/v1/games/multiget-place-details?placeIds={place_id}"
            try:
                async with self.get_session() as session:
                    async with session.get(details_url, timeout=5) as resp:
                        self.log_status(f"Place Details API returned HTTP {resp.status}")
                        if resp.status == 200:
                            data = await resp.json()
                            if isinstance(data, list) and len(data) > 0:
                                item = data[0]
                                results.append({
                                    "placeId": item.get("placeId", place_id),
                                    "name": item.get("name") or item.get("builder") or f"Place {place_id}"
                                })
                            elif isinstance(data, dict) and "data" in data and len(data["data"]) > 0:
                                item = data["data"][0]
                                results.append({
                                    "placeId": item.get("placeId", place_id),
                                    "name": item.get("name", f"Place {place_id}")
                                })
            except Exception as e:
                self.log_status(f"Error fetching place details: {e}", COLOR_ERROR)

            if not results:
                self.log_status(f"API gave no metadata for {place_id}, using fallback title.", COLOR_ACCENT)
                results.append({
                    "placeId": place_id,
                    "name": f"Roblox Game {place_id}"
                })
        else:
            self.log_status(f"Searching OmniSearch for query: '{query}'...", COLOR_ACCENT)
            url = f"https://apis.roblox.com/search-api/omni-search?searchQuery={query}&sessionId={self.session_id}&pageType=all"
            try:
                async with self.get_session() as session:
                    async with session.get(url, timeout=5) as resp:
                        if resp.status == 200:
                            data = await resp.json()
                            for result in data.get("searchResults", []):
                                for item in result.get("contents", []):
                                    if item.get("contentType") == "Game":
                                        results.append({
                                            "placeId": item.get("rootPlaceId"),
                                            "name": item.get("name")
                                        })
                            self.log_status(f"OmniSearch completed. Matches found: {len(results)}")
                        else:
                            self.log_status(f"OmniSearch API Error HTTP {resp.status}", COLOR_ERROR)
            except Exception as e:
                self.log_status(f"OmniSearch Exception: {e}", COLOR_ERROR)

        self.after(0, self._render_results, results)

        if results:
            self.async_task(self.load_thumbnails_sequential(results))

    def _render_results(self, results):
        self.search_button.configure(state="normal")

        for widget in self.results_scrollable.winfo_children():
            widget.destroy()

        self.cards_map.clear()

        if not results:
            self.log_status("No games found matching your query.", COLOR_ERROR)
            return

        self.log_status(f"Found {len(results)} game(s). Downloading icons...", COLOR_ACCENT)

        for idx, item in enumerate(results):
            place_id = item["placeId"]
            row = idx // 2
            col = idx % 2

            card = ctk.CTkFrame(
                self.results_scrollable,
                fg_color=COLOR_BG,
                border_color=COLOR_BORDER,
                border_width=1,
                corner_radius=14,
                height=110
            )
            card.grid(row=row, column=col, padx=6, pady=6, sticky="ew")
            card.grid_propagate(False)

            img_label = ctk.CTkLabel(
                card,
                text="Loading...",
                font=(FONT_FAMILY, 11),
                text_color=COLOR_TEXT_MUTED,
                width=86,
                height=86
            )
            img_label.pack(side="left", padx=12, pady=12)

            info_frame = ctk.CTkFrame(card, fg_color="transparent")
            info_frame.pack(side="left", fill="both", expand=True, padx=(0, 10), pady=10)

            title_img = render_text_with_color_emojis(item["name"], font_size=13)
            title_label = ctk.CTkLabel(info_frame, image=title_img, text="")
            title_label.pack(anchor="w")

            create_btn = ctk.CTkButton(
                info_frame,
                text="Create Shortcut",
                height=30,
                width=120,
                corner_radius=10,
                fg_color=COLOR_ACCENT,
                hover_color="#0284C7",
                text_color="#0B0F19",
                font=(FONT_FAMILY, 12, "bold"),
                command=lambda p=item: self.create_shortcut(p)
            )
            create_btn.pack(anchor="w", pady=(6, 0))

            self.cards_map[place_id] = {
                "card": card,
                "img_label": img_label,
                "item": item
            }

    async def load_thumbnails_sequential(self, items):
        async with self.get_session() as session:
            for item in items:
                place_id = item.get("placeId")
                if not place_id or place_id not in self.cards_map:
                    continue

                pil_img = await self._fetch_place_icon(session, place_id)
                if pil_img:
                    rounded_img = round_image_corners(pil_img, radius=14)
                    item["pil_image"] = pil_img
                    self.after(0, self._update_card_image, place_id, rounded_img)
                else:
                    self.after(0, self._update_card_error, place_id)

                await asyncio.sleep(0.05)

        self.log_status("All game icons loaded successfully.", COLOR_SUCCESS)

    async def _fetch_place_icon(self, session, place_id):
        api_url = f"https://thumbnails.roblox.com/v1/places/gameicons?placeIds={place_id}&returnPolicy=PlaceHolder&size=256x256&format=Png&isCircular=false"
        try:
            async with session.get(api_url, timeout=5) as resp:
                if resp.status == 200:
                    data = await resp.json()
                    if "data" in data and len(data["data"]) > 0:
                        image_url = data["data"][0].get("imageUrl")
                        if image_url:
                            async with session.get(image_url, timeout=5) as img_resp:
                                if img_resp.status == 200:
                                    img_bytes = await img_resp.read()
                                    return Image.open(io.BytesIO(img_bytes)).resize((86, 86))
        except Exception as e:
            self.log_status(f"Error fetching icon for place {place_id}: {e}", COLOR_ERROR)
        return None

    def _update_card_image(self, place_id, rounded_img):
        if place_id in self.cards_map:
            img_label = self.cards_map[place_id]["img_label"]
            ctk_img = ctk.CTkImage(light_image=rounded_img, dark_image=rounded_img, size=(86, 86))
            img_label.configure(image=ctk_img, text="")

    def _update_card_error(self, place_id):
        if place_id in self.cards_map:
            img_label = self.cards_map[place_id]["img_label"]
            img_label.configure(text="No Icon")

    def create_shortcut(self, item):
        place_id = item["placeId"]
        place_name = item["name"]
        pil_img = item.get("pil_image")

        user_profile = os.environ.get("USERPROFILE", os.path.expanduser("~"))
        desktop_path = os.path.join(user_profile, "Desktop")

        safe_name = "".join(c for c in place_name if c.isalnum() or c in (" ", "_", "-")).strip()
        if not safe_name:
            safe_name = f"Roblox_Game_{place_id}"

        shortcut_path = os.path.join(desktop_path, f"{safe_name}.lnk")

        real_appdata = os.path.join(user_profile, "AppData", "Roaming", "RobloxShortcuts", "Icons")
        os.makedirs(real_appdata, exist_ok=True)

        icon_path = os.path.join(real_appdata, f"{place_id}.ico")

        try:
            if pil_img:
                img_rgba = pil_img.convert("RGBA")
                img_rgba.save(
                    icon_path,
                    format="ICO",
                    sizes=[(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)]
                )

            pythoncom.CoInitialize()

            shell = win32com.client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortCut(shortcut_path)
            shortcut.TargetPath = f"roblox://placeid={place_id}"
            shortcut.WorkingDirectory = desktop_path

            if os.path.exists(icon_path):
                shortcut.IconLocation = f"{icon_path},0"

            shortcut.save()

            self.log_status(f"Shortcut '{safe_name}.lnk' created on Desktop!", COLOR_SUCCESS)
        except Exception as e:
            self.log_status(f"Failed to create shortcut: {e}", COLOR_ERROR)


if __name__ == "__main__":
    app = RobloxShortcutApp()
    app.mainloop()