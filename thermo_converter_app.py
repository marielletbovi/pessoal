import os
import sys
import shutil
import threading
import time
import datetime
import json
import re
import pandas as pd
import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import ctypes
from PIL import Image, ImageTk

# Force light appearance globally — prevents dark/black backgrounds
ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")

# --- Canva Fidedigno Color Palette ---
FUNDO_SIDEBAR = "#C5D6E5"        # Light steel blue sidebar background
AZUL_NAV_SELECTED = "#1A335E"    # Dark navy blue selected button
FUNDO_APP = "#EDF2F7"            # Soft light blue-gray background tint
FUNDO_SUPERFICIE = "#FFFFFF"    # Pure white card bodies
CRIMSON_HDR = "#B80000"          # Crimson red card headers and badges
YELLOW_BTN = "#ECA300"           # Yellow/gold directory buttons
ORANGE_BTN = "#E06F00"           # Orange manual convert button
BLUE_WATCH_BTN = "#7CAAD0"       # Slate blue watchdog button
GREEN_SUCCESS = "#00B050"        # Bright green success save button
TEXTO_NAV = "#1A335E"            # Dark navy text for deselected buttons
TEXTO = "#1B2430"                # Dark charcoal gray body text
TEXTO_SUAVE = "#4A5568"          # Slate gray descriptive text
BORDA = "#D1E0EC"                # Soft blue border color
GOLD_STRIP = "#C87000"           # Golden-orange bottom strip
CARD_RADIUS = 8
CONTROL_RADIUS = 8

FONT_HERO = ("Segoe UI Semibold", 20)
FONT_H1 = ("Segoe UI Semibold", 15)
FONT_H2 = ("Segoe UI Semibold", 12)
FONT_BODY = ("Segoe UI", 11)
FONT_BODY_BOLD = ("Segoe UI Semibold", 11)
FONT_CAPTION = ("Segoe UI", 9)

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and frozen apps """
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    base_dir = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, 'frozen', False) else __file__))
    path = os.path.join(base_dir, relative_path)
    if not os.path.exists(path):
        path = os.path.join(os.path.dirname(base_dir), relative_path)
    return path

# --- Configuration Persistence Manager ---
class ConfigManager:
    def __init__(self, config_path="config.json"):
        self.config_path = config_path
        self.data = self.get_defaults()
        self.load()

    def get_defaults(self):
        return {
            "pasta_origem": "",
            "pasta_backup": "",
            "pasta_destino": "",
            "zerar_negativos": True,
            "adicionar_respiro": False,
            "data_invertida": False,
            "incluir_hora_arquivo": True,
            "separador_campo": ";",
            "separador_decimal": ",",
            "casas_decimais": 4,
            "custom_filename_active": False,
            "custom_filename_mask": "[Sample Name]_[Timestamp da Conversão]",
            "colunas": [
                {"name": "ID#", "enabled": True, "source": "index"},
                {"name": "Sample Name", "enabled": True, "source": "metadata", "key": "sample_name"},
                {"name": "Usuário", "enabled": True, "source": "metadata", "key": "user"},
                {"name": "Data", "enabled": True, "source": "metadata", "key": "date"},
                {"name": "Método", "enabled": True, "source": "metadata", "key": "method"},
                {"name": "Dia", "enabled": True, "source": "metadata", "key": "day"},
                {"name": "Mês", "enabled": True, "source": "metadata", "key": "month"},
                {"name": "Hora", "enabled": True, "source": "metadata", "key": "hour"},
                {"name": "Original Files", "enabled": True, "source": "metadata", "key": "original_file"}
            ],
            "produtos": [
                {"name": "Sacarose", "original_names": ["Sacarose", "Sucrose"], "enabled": True},
                {"name": "Glicose", "original_names": ["Glicose", "Glucose"], "enabled": True},
                {"name": "Frutose", "original_names": ["Frutose", "Fructose"], "enabled": True},
                {"name": "Manitol", "original_names": ["Manitol", "Mannitol"], "enabled": True},
                {"name": "Glicerol", "original_names": ["Glicerol", "Glycerol"], "enabled": True},
                {"name": "Etanol", "original_names": ["Etanol", "Ethanol"], "enabled": True}
            ],
            "formulas": [
                {"name": "AR", "expression": "Glicose + Frutose", "enabled": True},
                {"name": "ART", "expression": "(Sacarose / 0.95) + Glicose + Frutose", "enabled": True}
            ],
            "extras": []
        }

    def load(self):
        if os.path.exists(self.config_path):
            try:
                with open(self.config_path, "r", encoding="utf-8") as f:
                    loaded_data = json.load(f)
                    for k, v in loaded_data.items():
                        self.data[k] = v
            except Exception as e:
                print(f"Error loading config.json: {e}")

    def save(self):
        try:
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(self.data, f, indent=4, ensure_ascii=False)
        except Exception as e:
            print(f"Error saving config.json: {e}")

    def reset(self):
        self.data = self.get_defaults()
        self.save()

# --- Core Excel Parser & Processor ("Batalha Naval") ---
class ThermoParser:
    def __init__(self, config_manager):
        self.config = config_manager

    def fix_bad_characters(self, text):
        if not isinstance(text, str):
            return text
        replacements = {
            "PULM\ufffdO": "PULMÃO",
            "BAGA\ufffdO": "BAGAÇO",
            "F\ufffdBRICA": "FÁBRICA",
            "EXTRA\ufffdO": "EXTRAÇÃO",
            "EVAPORA\ufffdO": "EVAPORAÇÃO",
            "SA\ufffdDA": "SAÍDA",
            "PR\ufffd": "PRÉ",
            "\ufffdGUA": "ÁGUA",
            "PULM\u00c3O": "PULMÃO",
            "BAGA\u00c7O": "BAGAÇO"
        }
        for k, v in replacements.items():
            text = text.replace(k, v)
        text = text.replace("\ufffd", "")
        return text

    def parse(self, filepath):
        if not os.path.exists(filepath):
            return None

        try:
            xl = pd.ExcelFile(filepath)
            if 'Integration' not in xl.sheet_names:
                return None
            df = pd.read_excel(xl, sheet_name='Integration', header=None)
        except Exception as e:
            print(f"Error reading {filepath}: {e}")
            return None

        layout = "Unknown"
        sample_name_raw = ""
        date_val = None
        user_val = "Unknown"
        method_val = "Unknown"

        # Check for Cruz Alta Layout
        if df.shape[1] >= 3 and str(df.iloc[1, 0]).strip().startswith('Sample Name'):
            layout = "Cruz Alta"
            sample_name_raw = str(df.iloc[1, 2]).strip()
            date_val = df.iloc[4, 2]
            method_val = str(df.iloc[3, 2]).strip()
            user_val = str(df.iloc[3, 6]).strip() if df.shape[1] > 6 else "Unknown"

        # Check for Tanabi Layout
        elif df.shape[1] >= 3 and (str(df.iloc[3, 0]).strip().startswith('Injection Name') or str(df.iloc[0, 0]).strip() == 'Chromatogram and Results'):
            layout = "Tanabi"
            sample_name_raw = str(df.iloc[3, 2]).strip()
            date_val = df.iloc[9, 2]
            method_val = str(df.iloc[7, 2]).strip()
            try:
                overview_df = pd.read_excel(xl, sheet_name='Overview', header=None)
                if overview_df.shape[0] > 6 and overview_df.shape[1] > 6 and not pd.isna(overview_df.iloc[6, 6]):
                    user_val = str(overview_df.iloc[6, 6]).strip()
                elif overview_df.shape[0] > 4 and overview_df.shape[1] > 6 and not pd.isna(overview_df.iloc[4, 6]):
                    user_val = str(overview_df.iloc[4, 6]).strip()
            except Exception:
                user_val = "Unknown"

        if layout == "Unknown":
            return None

        # Clean replacement character encodings
        sample_name_raw = self.fix_bad_characters(sample_name_raw)
        user_val = self.fix_bad_characters(user_val)
        method_val = self.fix_bad_characters(method_val)

        # Extract date parts
        date_str = ""
        day_str = ""
        month_str = ""
        
        if isinstance(date_val, datetime.datetime):
            date_str = date_val.strftime('%Y/%m/%d' if self.config.data.get("data_invertida", False) else '%d/%m/%Y')
            day_str = date_val.strftime('%d')
            month_str = date_val.strftime('%m')
        elif date_val is not None:
            date_str = str(date_val)
            date_match = re.search(r'(\d{2})/(\d{2})/(\d{4})', date_str)
            if date_match:
                day_str, month_str = date_match.group(1), date_match.group(2)
                if self.config.data.get("data_invertida", False):
                    date_str = f"{date_match.group(3)}/{month_str}/{day_str}"
            else:
                date_match_iso = re.search(r'(\d{4})-(\d{2})-(\d{2})', date_str)
                if date_match_iso:
                    day_str, month_str = date_match_iso.group(3), date_match_iso.group(2)
                    if not self.config.data.get("data_invertida", False):
                        date_str = f"{day_str}/{month_str}/{date_match_iso.group(1)}"

        # Extract nominal hour from sample name
        hour_val = ""
        hour_match = re.search(r'\b(\d{2}:\d{2})\b', sample_name_raw)
        if hour_match:
            hour_val = hour_match.group(1)
        else:
            if isinstance(date_val, datetime.datetime):
                hour_val = date_val.strftime('%H:%M')

        # Clean Sample Name (e.g. "Mosto - 08:00" -> "Mosto", "PULMÃO 08:00" -> "PULMÃO")
        sample_name_clean = re.sub(r'[-\s]*\b\d{2}:\d{2}\b', '', sample_name_raw).strip()

        # Parse Peak Results
        peaks = {}
        if layout == "Cruz Alta":
            start_row = 8
            for idx in range(start_row, df.shape[0]):
                p_name = df.iloc[idx, 2]
                p_amount = df.iloc[idx, 6]
                if pd.isna(p_name) or str(p_name).strip() == "" or str(p_name).strip().upper() == "TOTAL:":
                    continue
                peaks[str(p_name).strip().upper()] = p_amount
        elif layout == "Tanabi":
            start_row = 42
            for idx in range(start_row, df.shape[0]):
                p_name = df.iloc[idx, 1]
                p_amount = df.iloc[idx, 7]
                if pd.isna(p_name) or str(p_name).strip() == "" or str(p_name).strip().upper() == "TOTAL:":
                    continue
                peaks[str(p_name).strip().upper()] = p_amount

        # Build Metadata Context
        metadata = {
            "sample_name": sample_name_clean,
            "sample_name_raw": sample_name_raw,
            "user": user_val,
            "date": date_str,
            "method": method_val,
            "day": day_str,
            "month": month_str,
            "hour": hour_val,
            "original_file": os.path.basename(filepath),
            "layout": layout
        }

        # Build Product Concentrations
        concentrations = {}
        zerar = self.config.data.get("zerar_negativos", True)

        for prod in self.config.data["produtos"]:
            name = prod["name"]
            aliases = [name.upper()] + [a.upper() for a in prod.get("original_names", [])]
            
            val = 0.0
            for alias in aliases:
                if alias in peaks:
                    raw_val = peaks[alias]
                    try:
                        val = float(str(raw_val).replace(",", "."))
                    except ValueError:
                        val = 0.0
                    break
            
            if zerar and val < 0:
                val = 0.0
                
            concentrations[name] = val

        return {
            "metadata": metadata,
            "concentrations": concentrations
        }

    def evaluate_formulas(self, concentrations):
        results = {}
        zerar = self.config.data.get("zerar_negativos", True)
        
        vals = {k: float(v) for k, v in concentrations.items()}
        if zerar:
            for k in vals:
                if vals[k] < 0:
                    vals[k] = 0.0
                    
        for formula in self.config.data["formulas"]:
            if not formula.get("enabled", True):
                continue
            name = formula["name"]
            expr = formula["expression"]
            
            sorted_keys = sorted(vals.keys(), key=len, reverse=True)
            eval_expr = expr
            for k in sorted_keys:
                eval_expr = re.sub(rf'\b{k}\b', str(vals[k]), eval_expr)
                
            if not re.match(r'^[0-9+\-*/().\s]*$', eval_expr):
                results[name] = 0.0
                continue
                
            try:
                val = eval(eval_expr)
                if pd.isna(val):
                    val = 0.0
                if val < 0 and zerar:
                    val = 0.0
                results[name] = round(val, 6)
            except Exception:
                results[name] = 0.0
                
        # Also handle extra columns empty values
        for ext in self.config.data.get("extras", []):
            if ext.get("enabled", True):
                results[ext["name"]] = 0.0
                
        return results

    def get_decimal_places(self):
        try:
            return max(0, min(8, int(self.config.data.get("casas_decimais", 2))))
        except (TypeError, ValueError):
            return 2

    def build_row(self, data):
        calc_vals = self.evaluate_formulas(data["concentrations"])
        
        row = {}
        sep_dec = self.config.data.get("separador_decimal", ",")
        decimal_places = self.get_decimal_places()
        
        def format_val(v, is_numeric=False):
            if is_numeric:
                try:
                    f = float(v)
                    return f"{f:.{decimal_places}f}".replace(".", sep_dec)
                except ValueError:
                    return f"{0:.{decimal_places}f}".replace(".", sep_dec)
            s = str(v)
            if re.match(r'^-?\d+\.?\d*$', s.replace(',', '.')):
                return s.replace('.', sep_dec).replace(',', sep_dec)
            return s

        # Build Context Map
        context_map = {}
        for k, v in data["metadata"].items():
            context_map[k] = v
        for k, v in data["concentrations"].items():
            context_map[k] = v
        for k, v in calc_vals.items():
            context_map[k] = v

        # Columns order from config list
        idx_enabled = False
        idx_name = "ID#"
        enabled_items = []

        # Read the unified order of columns as defined by metadata order, products order, and formulas
        for col in self.config.data["colunas"]:
            if col["name"] == "ID#":
                idx_enabled = col.get("enabled", True)
                idx_name = col.get("name", "ID#")
                continue
            if col.get("enabled", True):
                enabled_items.append((col["name"], col.get("key"), "metadata"))

        for prod in self.config.data["produtos"]:
            if prod.get("enabled", True):
                enabled_items.append((prod["name"], prod["name"], "product"))

        for formula in self.config.data["formulas"]:
            if formula.get("enabled", True):
                enabled_items.append((formula["name"], formula["name"], "formula"))

        for ext in self.config.data.get("extras", []):
            if ext.get("enabled", True):
                enabled_items.append((ext["name"], ext["name"], "extra"))

        if idx_enabled:
            row[idx_name] = "1"

        for col_title, key, col_type in enabled_items:
            val = context_map.get(key, "")
            is_num = (col_type in ["product", "formula", "extra"])
            row[col_title] = format_val(val, is_numeric=is_num)

        return row, data["metadata"]["sample_name"]

    def convert_file(self, filepath, dest_folder):
        parsed = self.parse(filepath)
        if not parsed:
            return False

        row_data, sample_name_clean = self.build_row(parsed)
        
        # Format filename
        now = datetime.datetime.now()
        
        # Resolve date formatting
        if self.config.data.get("data_invertida", False):
            date_str = now.strftime("%Y%m%d")
        else:
            date_str = now.strftime("%d%m%Y")
            
        # Resolve hour inclusion
        hour_str = now.strftime("%H%M") if self.config.data.get("incluir_hora_arquivo", True) else ""
        
        ts = f"{date_str}{hour_str}"
        
        safe_name = re.sub(r'[\\/*?:"<>|]', "-", sample_name_clean)
        if not safe_name:
            safe_name = "Amostra"
            
        if self.config.data.get("custom_filename_active", False):
            mask = self.config.data.get("custom_filename_mask", "[Sample Name]_[Timestamp da Conversão]")
            mask_resolved = self.resolve_filename_template(mask, parsed, row_data, now)
            out_name = f"{mask_resolved}.csv"
        else:
            out_name = f"{safe_name}_{ts}.csv"
            
        out_path = os.path.join(dest_folder, out_name)
        
        # Handle duplicates safely
        counter = 1
        base, ext = os.path.splitext(out_path)
        while os.path.exists(out_path):
            out_path = f"{base} ({counter}){ext}"
            counter += 1

        final_list = []
        respiro = self.config.data.get("adicionar_respiro", True)
        if respiro:
            final_list.append({k: "" for k in row_data.keys()})
            
        final_list.append(row_data)
        
        df = pd.DataFrame(final_list)
        sep_campo = self.config.data.get("separador_campo", ";")
        
        try:
            df.to_csv(out_path, index=False, sep=sep_campo, encoding="latin1")
            return True
        except Exception as e:
            print(f"Error saving CSV: {e}")
            return False

    def resolve_filename_template(self, template, parsed, row_data, now):
        metadata = parsed.get("metadata", {})
        tokens = {
            "Sample Name": metadata.get("sample_name", ""),
            "Amostra": metadata.get("sample_name", ""),
            "Data do Arquivo": metadata.get("date", ""),
            "Data da Conversão": now.strftime("%d/%m/%Y"),
            "Hora da Conversão": now.strftime("%H:%M"),
            "Timestamp da Conversão": now.strftime("%Y%m%d%H%M"),
            "Timestamp ddmmaaaahhmm": now.strftime("%d%m%Y%H%M"),
            "Original Files": metadata.get("original_file", ""),
            "Original File": metadata.get("original_file", ""),
            "Layout": metadata.get("layout", "")
        }
        tokens.update({k: v for k, v in row_data.items()})

        def replace_token(match):
            key = match.group(1).strip()
            return str(tokens.get(key, ""))

        resolved = re.sub(r"\[([^\]]+)\]", replace_token, template or "")
        resolved = re.sub(r"'([^']*)'", r"\1", resolved)
        resolved = resolved.strip() or tokens.get("Sample Name") or "Amostra"
        resolved = re.sub(r'[\\/*?:"<>|]', "-", resolved)
        resolved = re.sub(r"\s+", " ", resolved).strip()
        return resolved or "Amostra"

# --- Custom Settings Pop-up Frame/Window for Gear Icons ---
class GearConfigDialog(ctk.CTkToplevel):
    def __init__(self, parent, item, item_type, config_manager, save_callback):
        super().__init__(parent)
        self.config_manager = config_manager
        self.item = item
        self.item_type = item_type
        self.save_callback = save_callback
        
        self.title(f"Ajuste - {item['name']}")
        self.geometry("450x300")
        self.configure(fg_color=FUNDO_APP)
        self.grab_set()
        
        # Theme icon
        self.after(200, self.apply_icon)
        
        header = ctk.CTkFrame(self, height= 50, fg_color=CRIMSON_HDR, corner_radius=0)
        header.pack(fill="x")
        ctk.CTkLabel(header, text=f"⚙️ Configuração: {item['name']}", font=("Segoe UI", 12, "bold"), text_color="white").pack(side="left", padx=20)

        body = ctk.CTkFrame(self, fg_color=FUNDO_SUPERFICIE, corner_radius=12, border_width=1, border_color=BORDA)
        body.pack(fill="both", expand=True, padx=15, pady=15)

        # Name label and edit
        ctk.CTkLabel(body, text="Nome de Exibição Final no CSV:", font=FONT_BODY_BOLD, text_color=TEXTO).pack(anchor="w", padx=20, pady=(15, 2))
        self.ent_name = ctk.CTkEntry(body, width=320,
                                     fg_color="white", border_color=BORDA, text_color=TEXTO)
        self.ent_name.pack(padx=20, pady=(0, 15))
        self.ent_name.insert(0, item["name"])

        # Aliases edit (only for products)
        self.ent_aliases = None
        if item_type == "product":
            ctk.CTkLabel(body, text="Nomes originais no cromatograma (separados por vírgula):", font=FONT_BODY_BOLD, text_color=TEXTO).pack(anchor="w", padx=20, pady=(0, 2))
            self.ent_aliases = ctk.CTkEntry(body, width=320,
                                            fg_color="white", border_color=BORDA, text_color=TEXTO)
            self.ent_aliases.pack(padx=20, pady=(0, 15))
            self.ent_aliases.insert(0, ", ".join(item.get("original_names", [])))

        # Bottom save
        btn_save = ctk.CTkButton(self, text="SALVAR AJUSTES", fg_color=GREEN_SUCCESS, hover_color="#218838",
                                 text_color="white", font=FONT_BODY_BOLD, height=35, command=self.save_data)
        btn_save.pack(pady=(0, 15))

    def apply_icon(self):
        icon_path = resource_path("logo.ico")
        if os.path.exists(icon_path):
            try: self.iconbitmap(icon_path)
            except: pass

    def save_data(self):
        new_name = self.ent_name.get().strip()
        if not new_name:
            messagebox.showwarning("Aviso", "Nome não pode ser vazio!")
            return
            
        self.item["name"] = new_name
        if self.item_type == "product" and self.ent_aliases:
            self.item["original_names"] = [x.strip() for x in self.ent_aliases.get().split(",") if x.strip()]
            
        self.config_manager.save()
        self.save_callback()
        self.destroy()

# --- Main Application Graphical Interface ---
class ThermoConverterApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.config_manager = ConfigManager()
        self.parser = ThermoParser(self.config_manager)
        
        self.title("Thermo Chromatography Converter")
        self.geometry("1100x720")
        self.minsize(720, 520)
        self.configure(fg_color=FUNDO_APP)
        self.compact_mode = False
        
        # Set taskbar icon explicit ID
        try:
            myappid = 'thermo.conversor.v2.Canva'
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        except:
            pass

        self.apply_theme_icon()
        
        # Watchdog Sentinel state
        self.sentinel_active = False
        self.sentinel_thread = None
        self.converted_count = 0

        self.setup_layout()
        self.bind("<Configure>", self.on_window_configure)
        
        # Navigate to home by default
        self.select_page("home")

    def apply_theme_icon(self):
        icon_path = resource_path("logo.ico")
        if os.path.exists(icon_path):
            try: self.iconbitmap(icon_path)
            except: pass

    def setup_layout(self):
        # Sidebar Left Frame
        self.sidebar = ctk.CTkFrame(self, width=220, fg_color=FUNDO_SIDEBAR, corner_radius=0)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        # Brand / Logo inside Sidebar
        logo_path = resource_path("logo.png")
        if os.path.exists(logo_path):
            try:
                img = Image.open(logo_path)
                max_w, max_h =170, 76
                ratio = min(max_w / img.width, max_h / img.height)
                logo_size = (max(1, int(img.width * ratio)), max(1, int(img.height * ratio)))
                self.logo_img = ctk.CTkImage(light_image=img, dark_image=img, size=logo_size)
                compact_ratio = min(52 / img.width, 56 / img.height)
                compact_size = (max(1, int(img.width * compact_ratio)), max(1, int(img.height * compact_ratio)))
                self.logo_img_compact = ctk.CTkImage(light_image=img, dark_image=img, size=compact_size)
                self.logo_lbl = ctk.CTkLabel(self.sidebar, image=self.logo_img, text="")
                self.logo_lbl.pack(pady=(18, 16), padx=20)
            except Exception:
                ctk.CTkLabel(self.sidebar, text="TEREOS", font=("Segoe UI", 24, "bold"), text_color=AZUL_NAV_SELECTED).pack(pady=20)
        else:
            ctk.CTkLabel(self.sidebar, text="TEREOS", font=("Segoe UI", 24, "bold"), text_color=AZUL_NAV_SELECTED).pack(pady=20)

        # Sidebar Buttons
        self.nav_buttons = {}
        
        # Page link: Início
        self.create_nav_button("home", "Início", self.sidebar, is_main=True)
        
        # Divider line
        ctk.CTkFrame(self.sidebar, height=1, fg_color="white").pack(fill="x", padx=15, pady=8)
        
        # "Configurações" bold category
        lbl_cat = ctk.CTkLabel(self.sidebar, text="Configurações", font=("Segoe UI Semibold", 12, "bold"), text_color=AZUL_NAV_SELECTED)
        lbl_cat.pack(anchor="w", padx=25, pady=(2, 4))
        
        # Config options
        self.create_nav_button("general", "Gerais", self.sidebar)
        self.create_nav_button("columns", "Colunas", self.sidebar)
        
        # Visualização option
        ctk.CTkFrame(self.sidebar, height=1, fg_color="white").pack(fill="x", padx=15, pady=8)
        self.create_nav_button("preview", "Visualização", self.sidebar)

        # Sidebar Footer elements (Status, Reset, Help)
        self.sidebar_footer = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        self.sidebar_footer.place(relx=0.5, rely=0.76, anchor="center", relwidth=0.86)
        
        # STATUS badge
        status_hdr = ctk.CTkFrame(self.sidebar_footer, fg_color=CRIMSON_HDR, height=28, corner_radius=CONTROL_RADIUS)
        status_hdr.pack(fill="x", pady=2)
        status_hdr.pack_propagate(False)
        ctk.CTkLabel(status_hdr, text="STATUS", font=("Segoe UI", 10, "bold"), text_color="white").pack(expand=True)

        # Status text box (styled as card white with rounded corners and light border)
        self.status_box = ctk.CTkFrame(self.sidebar_footer, fg_color="white", corner_radius=CARD_RADIUS, border_width=1, border_color=BORDA, height=64)
        self.status_box.pack(fill="x", pady=2)
        self.status_box.pack_propagate(False)
        
        self.status_box.grid_columnconfigure(0, weight=1)
        self.status_box.grid_columnconfigure(1, weight=1)
        self.status_box.grid_rowconfigure(0, weight=1)
        self.status_box.grid_rowconfigure(1, weight=1)

        self.lbl_arquivos_label = ctk.CTkLabel(self.status_box, text="Arquivos", font=FONT_BODY, text_color=TEXTO_SUAVE, anchor="w")
        self.lbl_arquivos_label.grid(row=0, column=0, padx=(12, 0), pady=(4, 0), sticky="w")

        self.lbl_status_n = ctk.CTkLabel(self.status_box, text="0", font=FONT_BODY_BOLD, text_color=TEXTO, anchor="e")
        self.lbl_status_n.grid(row=0, column=1, padx=(0, 12), pady=(4, 0), sticky="e")

        self.lbl_napasta_label = ctk.CTkLabel(self.status_box, text="na Pasta:", font=FONT_BODY, text_color=TEXTO_SUAVE, anchor="w")
        self.lbl_napasta_label.grid(row=1, column=0, padx=(12, 0), pady=(0, 4), sticky="w")

        self.lbl_status_arquivos = ctk.CTkLabel(self.status_box, text="arquivos", font=FONT_BODY_BOLD, text_color=TEXTO, anchor="e")
        self.lbl_status_arquivos.grid(row=1, column=1, padx=(0, 12), pady=(0, 4), sticky="e")

        # Restore default button (navy color, wrapped 3 lines)
        btn_reset = ctk.CTkButton(self.sidebar_footer, text="RESTAURAR\nCONFIGURAÇÃO\nPADRÃO", height=54, 
                                  fg_color=AZUL_NAV_SELECTED, hover_color="#2E79B9", text_color="white", font=("Segoe UI", 9, "bold"),
                                  corner_radius=CONTROL_RADIUS, command=self.reset_defaults)
        btn_reset.pack(fill="x", pady=(6, 2))

        # Help round button (aligned right)
        footer_btn_frame = ctk.CTkFrame(self.sidebar_footer, fg_color="transparent")
        footer_btn_frame.pack(fill="x", pady=(4, 0))
        btn_help = ctk.CTkButton(footer_btn_frame, text="?", width=28, height=28, 
                                 fg_color="#86B3D9", hover_color="#6393C0", text_color="white", font=("Segoe UI", 12, "bold"),
                                 corner_radius=14, command=self.show_help)
        btn_help.pack(side="right")

        # Golden bottom strip (accent bar inside window footer or sidebar)
        self.bottom_strip = ctk.CTkFrame(self, height=4, fg_color=GOLD_STRIP, corner_radius=0)
        self.bottom_strip.pack(fill="x", side="bottom")

        # Container for main page frames
        self.content_container = ctk.CTkFrame(self, fg_color="transparent", corner_radius=0)
        self.content_container.pack(side="right", fill="both", expand=True)

        # Page Frames Initialization
        self.frames = {
            "home": self.create_home_frame(),
            "general": self.create_general_frame(),
            "columns": self.create_columns_frame(),
            "preview": self.create_preview_frame()
        }

    def on_window_configure(self, event):
        if event.widget is not self:
            return

        compact = event.width < 920 or event.height < 640
        if compact == self.compact_mode:
            return

        self.compact_mode = compact
        if compact:
            self.sidebar.configure(width=178)
            for btn in self.nav_buttons.values():
                btn.configure(width=148, height=31, font=("Segoe UI", 10))
            if hasattr(self, "logo_img_compact"):
                self.logo_lbl.configure(image=self.logo_img_compact)
            self.sidebar_footer.place_configure(rely=0.78, relwidth=0.88)
            self.status_box.configure(height=54)
        else:
            self.sidebar.configure(width=220)
            for key, btn in self.nav_buttons.items():
                btn.configure(width=190, height=35, font=("Segoe UI", 12, "bold") if key == "home" else ("Segoe UI", 12))
            if hasattr(self, "logo_img"):
                self.logo_lbl.configure(image=self.logo_img)
            self.sidebar_footer.place_configure(rely=0.80, relwidth=0.86)
            self.status_box.configure(height=64)

        if hasattr(self, "current_page_key"):
            pad = 12 if compact else 24
            self.frames[self.current_page_key].pack_configure(padx=pad, pady=pad)

    def create_nav_button(self, key, label, parent, is_main=False):
        btn = ctk.CTkButton(parent, text=label, height=35, width=190,
                            fg_color="transparent", text_color=TEXTO_NAV,
                            font=("Segoe UI", 12, "bold") if is_main else ("Segoe UI", 12), anchor="w",
                            hover_color="#B5C9DC", corner_radius=CONTROL_RADIUS,
                            command=lambda k=key: self.select_page(k))
        btn.pack(pady=2, padx=15)
        self.nav_buttons[key] = btn

    def select_page(self, key):
        self.current_page_key = key

        # Deselect navigation buttons
        for k, btn in self.nav_buttons.items():
            btn.configure(fg_color="transparent", text_color=TEXTO_NAV)
            
        # Select active navigation button (Pill style dark blue)
        self.nav_buttons[key].configure(fg_color=AZUL_NAV_SELECTED, text_color="white")

        # Hide all frames
        for frame in self.frames.values():
            frame.pack_forget()

        # Show selected frame
        pad = 12 if self.compact_mode else 24
        self.frames[key].pack(fill="both", expand=True, padx=pad, pady=pad)
        
        # Specific page refreshes
        if key == "preview":
            self.refresh_preview()
        elif key == "columns":
            self.refresh_columns_lists()
            
        self.update_folder_counts()

    # --- SCREEN 1: HOME PAGE ---
    def create_home_frame(self):
        frame = ctk.CTkFrame(self.content_container, fg_color="transparent")
        
        # CARD 1 - CONFIGURAÇÃO DE DIRETÓRIOS
        card_dir = ctk.CTkFrame(frame, fg_color=FUNDO_SUPERFICIE, corner_radius=CARD_RADIUS, border_width=1, border_color=BORDA)
        card_dir.pack(fill="x", pady=(0, 20))
        
        # Header Crimson Bar
        card_dir_hdr = ctk.CTkFrame(card_dir, fg_color=CRIMSON_HDR, height=40, corner_radius=CONTROL_RADIUS)
        card_dir_hdr.pack(fill="x", padx=10, pady=10)
        card_dir_hdr.pack_propagate(False)
        ctk.CTkLabel(card_dir_hdr, text="CONFIGURAÇÃO DE DIRETÓRIOS", font=("Segoe UI", 11, "bold"), text_color="white").pack(expand=True)

        self.dir_labels = {}
        self.create_dir_selector(card_dir, "Pasta de Origem", "pasta_origem")
        self.create_dir_selector(card_dir, "Pasta de Backup", "pasta_backup")
        self.create_dir_selector(card_dir, "Pasta de Destino", "pasta_destino")
        
        # Spacer bottom inside card
        ctk.CTkLabel(card_dir, text="", height=5).pack()

        # LOWER SECTION - CONVERT AND WATCHDOG PANELS (2 COLUMNS)
        lower_frame = ctk.CTkFrame(frame, fg_color="transparent")
        lower_frame.pack(fill="both", expand=True)
        
        lower_frame.grid_columnconfigure(0, weight=1)
        lower_frame.grid_columnconfigure(1, weight=0)
        lower_frame.grid_columnconfigure(2, weight=1)
        lower_frame.grid_rowconfigure(0, weight=1)

        # Left Column - Manual Convert Panel (Transparent to sit on the blue-gray background)
        col_left = ctk.CTkFrame(lower_frame, fg_color="transparent")
        col_left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        
        # Text Header
        ctk.CTkLabel(col_left, text="CONVERTER ARQUIVOS\nDA PASTA DE ORIGEM", font=("Segoe UI", 14, "bold"), text_color=TEXTO).pack(pady=(25, 20))
        
        # Big Play Button Orange (Pill shape)
        self.btn_manual_convert = ctk.CTkButton(col_left, text="▶", font=("Segoe UI", 26, "bold"), width=180, height=65,
                                                 fg_color=ORANGE_BTN, hover_color="#C05C00", text_color="white",
                                                 corner_radius=32, command=self.run_manual_conversion)
        self.btn_manual_convert.pack(pady=15)

        # Vertical divider line between panels (Canva design line)
        divider = ctk.CTkFrame(lower_frame, width=2, fg_color="white")
        divider.grid(row=0, column=1, sticky="ns", padx=10, pady=20)

        # Right Column - Sentinel Watchdog Panel (Transparent to sit on the blue-gray background)
        col_right = ctk.CTkFrame(lower_frame, fg_color="transparent")
        col_right.grid(row=0, column=2, sticky="nsew", padx=(10, 0))

        # Text Header
        ctk.CTkLabel(col_right, text="LIGAR MONITORAMENTO\nDA PASTA", font=("Segoe UI", 14, "bold"), text_color=TEXTO).pack(pady=(25, 20))

        # Big Play Button Blue (Pill shape)
        self.btn_toggle_sentinel = ctk.CTkButton(col_right, text="▶", font=("Segoe UI", 26, "bold"), width=180, height=65,
                                                 fg_color=BLUE_WATCH_BTN, hover_color="#638CBA", text_color="white",
                                                 corner_radius=32, command=self.toggle_sentinel)
        self.btn_toggle_sentinel.pack(pady=15)

        # Ativo indicator
        self.lbl_sentinel_indicator = ctk.CTkLabel(col_right, text="✔ Ativo | 0 convertidos", font=("Segoe UI", 11, "bold"), text_color=GREEN_SUCCESS)
        self.lbl_sentinel_indicator.pack(pady=(2, 0))
        self.lbl_sentinel_indicator.pack_forget()  # Hidden by default, shown when active

        # Progress elements in home frame
        self.progress_lbl = ctk.CTkLabel(frame, text="", font=FONT_BODY, text_color=TEXTO_SUAVE)
        self.progress_lbl.pack(anchor="w", pady=(10, 2))
        self.progress_bar = ctk.CTkProgressBar(frame, height=8, fg_color=BORDA, progress_color=AZUL_NAV_SELECTED)
        self.progress_bar.pack(fill="x")
        self.progress_bar.set(0)

        return frame

    def create_dir_selector(self, parent, label, config_key):
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(fill="x", padx=30, pady=6)
        
        # Yellow button styled as perfect pill shape
        btn = ctk.CTkButton(frame, text=label, width=140, height=35,
                            fg_color=YELLOW_BTN, hover_color="#C98A00",
                            font=("Segoe UI", 10, "bold"), text_color="black", corner_radius=17,
                            command=lambda k=config_key: self.browse_folder(k))
        btn.pack(side="left")
        
        path = self.config_manager.data.get(config_key, "")
        lbl = ctk.CTkLabel(frame, text=path if path else "Diretório completo",
                           font=FONT_BODY, text_color=TEXTO_SUAVE, anchor="w")
        lbl.pack(side="left", expand=True, fill="x", padx=(30, 0))
        self.dir_labels[config_key] = lbl

    def browse_folder(self, key):
        path = filedialog.askdirectory()
        if path:
            full_path = os.path.abspath(path)
            self.config_manager.data[key] = full_path
            self.config_manager.save()
            self.dir_labels[key].configure(text=full_path)
            self.update_folder_counts()

    def update_folder_counts(self):
        orig = self.config_manager.data.get("pasta_origem", "")
        count = 0
        if orig and os.path.exists(orig):
            count = len([f for f in os.listdir(orig) if f.lower().endswith(('.xls', '.xlsx'))])
        
        if hasattr(self, "lbl_status_n"):
            self.lbl_status_n.configure(text=str(count))
        elif hasattr(self, "lbl_status_files"):
            self.lbl_status_files.configure(text=f"Arquivos\nna Pasta: {count} arquivos")
            
        if self.sentinel_active:
            self.lbl_sentinel_indicator.configure(text=f"✔ Ativo | {self.converted_count} convertidos")

    # --- SCREEN 2: GENERAL CONFIGS PAGE ---
    def create_general_frame(self):
        frame = ctk.CTkFrame(self.content_container, fg_color="transparent")
        
        # Main card
        card = ctk.CTkFrame(frame,
                            fg_color=FUNDO_SUPERFICIE,
                            corner_radius=CARD_RADIUS, border_width=1, border_color=BORDA)
        card.pack(fill="both", expand=True, pady=(0, 15))

        # CARD 1 HEADER - CONFIGURAÇÕES GERAIS
        card_dir_hdr = ctk.CTkFrame(card, fg_color=CRIMSON_HDR, height=40, corner_radius=CONTROL_RADIUS)
        card_dir_hdr.pack(fill="x", padx=10, pady=10)
        card_dir_hdr.pack_propagate(False)
        ctk.CTkLabel(card_dir_hdr, text="CONFIGURAÇÕES GERAIS", font=("Segoe UI", 11, "bold"), text_color="white").pack(expand=True)

        # Scrollable inner area
        scroll_card = ctk.CTkScrollableFrame(card,
                                             fg_color="transparent",
                                             scrollbar_button_color=BORDA,
                                             scrollbar_button_hover_color="#90A4BC")
        scroll_card.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        # Checkbox elements styled with custom X
        self.chk_respiro = self.create_canva_checkbox(scroll_card, "Adicionar linha de respiro", "adicionar_respiro")
        self.chk_zerar = self.create_canva_checkbox(scroll_card, "Zerar negativos", "zerar_negativos")

        # Custom File Name badge and options
        lbl_fn_badge = ctk.CTkButton(scroll_card, text="Nome do Arquivo final", width=160, height=25, 
                                     fg_color=YELLOW_BTN, hover_color=YELLOW_BTN, text_color="black", 
                                     text_color_disabled="black",
                                     font=("Segoe UI", 9, "bold"), corner_radius=CONTROL_RADIUS, state="disabled")
        lbl_fn_badge.pack(anchor="w", padx=25, pady=(15, 5))

        # Filename choices
        self.chk_fn_default = self.create_canva_checkbox(scroll_card, "Padrão", "custom_filename_active", invert_checked=True)
        
        fn_custom_frame = ctk.CTkFrame(scroll_card, fg_color="transparent")
        fn_custom_frame.pack(fill="x", padx=25, pady=2)

        self.chk_fn_custom = self.create_canva_checkbox(fn_custom_frame, "Personalizar", "custom_filename_active", pack_it=False)
        self.chk_fn_custom.pack(side="left")

        self.ent_fn_mask = ctk.CTkEntry(fn_custom_frame, height=28,
                                        fg_color="white", border_color=BORDA,
                                        text_color=TEXTO,
                                         corner_radius=8, font=FONT_CAPTION)
        self.ent_fn_mask.pack(side="left", fill="x", expand=True, padx=(15, 0))
        self.ent_fn_mask.insert(0, self.config_manager.data.get("custom_filename_mask", "[Sample Name]_[Timestamp da Conversão]"))
        self.ent_fn_mask.bind("<KeyRelease>", self.on_filename_template_keyrelease)
        self.chk_fn_default.configure(command=self.select_default_filename)
        self.chk_fn_custom.configure(command=self.select_custom_filename)

        # Horizontal layout for Data Invertida and Hora checkboxes
        dt_frame = ctk.CTkFrame(scroll_card, fg_color="transparent")
        dt_frame.pack(fill="x", padx=25, pady=5)
        self.chk_data_inv = self.create_canva_checkbox(dt_frame, "Data invertida (AAAA/MM/DD)", "data_invertida", pack_it=False)
        self.chk_data_inv.pack(side="left")
        self.chk_hora = self.create_canva_checkbox(dt_frame, "Hora", "incluir_hora_arquivo", pack_it=False)
        self.chk_hora.pack(side="left", padx=(30, 0))

        # Separator selectors
        sep_badge = ctk.CTkButton(scroll_card, text="Separadores CSV", width=160, height=25,
                                  fg_color=YELLOW_BTN, hover_color=YELLOW_BTN, text_color="black",
                                  text_color_disabled="black",
                                  font=("Segoe UI", 9, "bold"), corner_radius=CONTROL_RADIUS, state="disabled")
        sep_badge.pack(anchor="w", padx=25, pady=(15, 5))

        sep_frame = ctk.CTkFrame(scroll_card, fg_color="transparent")
        sep_frame.pack(fill="x", padx=25, pady=4)

        ctk.CTkLabel(sep_frame, text="Campo:", font=FONT_BODY_BOLD, text_color=TEXTO).pack(side="left")
        self.opt_sep_campo = ctk.CTkOptionMenu(
            sep_frame, values=[";", ",", "|", "\\t"],
            font=FONT_BODY, width=90,
            fg_color="white", button_color=BORDA, button_hover_color="#9EBDD8",
            text_color=TEXTO, dropdown_fg_color="white",
            dropdown_text_color=TEXTO, dropdown_hover_color="#9EBDD8")
        self.opt_sep_campo.pack(side="left", padx=(8, 30))
        self.opt_sep_campo.set(self.config_manager.data.get("separador_campo", ";"))

        ctk.CTkLabel(sep_frame, text="Decimal:", font=FONT_BODY_BOLD, text_color=TEXTO).pack(side="left")
        self.opt_sep_dec = ctk.CTkOptionMenu(
            sep_frame, values=[",", "."],
            font=FONT_BODY, width=90,
            fg_color="white", button_color=BORDA, button_hover_color="#C5D6E5",
            text_color=TEXTO, dropdown_fg_color="white",
            dropdown_text_color=TEXTO, dropdown_hover_color="#EDF2F7")
        self.opt_sep_dec.pack(side="left", padx=(8, 0))
        self.opt_sep_dec.set(self.config_manager.data.get("separador_decimal", ","))

        ctk.CTkLabel(sep_frame, text="Casas:", font=FONT_BODY_BOLD, text_color=TEXTO).pack(side="left", padx=(30, 0))
        self.opt_dec_places = ctk.CTkOptionMenu(
            sep_frame, values=["0", "1", "2", "3", "4", "5", "6"],
            font=FONT_BODY, width=80,
            fg_color="white", button_color=BORDA, button_hover_color="#C5D6E5",
            text_color=TEXTO, dropdown_fg_color="white",
            dropdown_text_color=TEXTO, dropdown_hover_color="#EDF2F7")
        self.opt_dec_places.pack(side="left", padx=(8, 0))
        self.opt_dec_places.set(str(self.config_manager.data.get("casas_decimais", 2)))

        # FÓRMULAS CARD INSIDE SCROLLABLE AREA
        card_formula_hdr = ctk.CTkFrame(scroll_card, fg_color=CRIMSON_HDR, height=40, corner_radius=CONTROL_RADIUS)
        card_formula_hdr.pack(fill="x", padx=10, pady=(25, 10))
        card_formula_hdr.pack_propagate(False)
        ctk.CTkLabel(card_formula_hdr, text="FÓRMULAS", font=("Segoe UI", 11, "bold"), text_color="white").pack(expand=True)

        # Headers for Coluna and Fórmula
        col_hdr_frame = ctk.CTkFrame(scroll_card, fg_color="transparent")
        col_hdr_frame.pack(fill="x", padx=25, pady=2)
        
        col_lbl1 = ctk.CTkButton(col_hdr_frame, text="Coluna", width=120, height=25, fg_color="#D1E0EC", text_color=TEXTO_NAV, text_color_disabled=TEXTO_NAV, font=FONT_BODY_BOLD, state="disabled")
        col_lbl1.pack(side="left")
        
        col_lbl2 = ctk.CTkButton(col_hdr_frame, text="Fórmula", width=120, height=25, fg_color="#D1E0EC", text_color=TEXTO_NAV, text_color_disabled=TEXTO_NAV, font=FONT_BODY_BOLD, state="disabled")
        col_lbl2.pack(side="left", padx=(15, 0))

        # Dynamic formulas rows
        self.formula_rows_container = ctk.CTkFrame(scroll_card, fg_color="transparent")
        self.formula_rows_container.pack(fill="x", padx=25, pady=5)
        self.draw_formulas_editor_rows()

        # Save Button bottom right
        btn_save = ctk.CTkButton(frame, text="SALVAR", height=40, width=150,
                                 fg_color=GREEN_SUCCESS, hover_color="#218838",
                                 font=("Segoe UI", 13, "bold"), command=self.save_general_settings,
                                 corner_radius=CONTROL_RADIUS)
        btn_save.pack(side="right")

        return frame

    def select_default_filename(self):
        self.chk_fn_default.select()
        self.chk_fn_custom.deselect()
        self.config_manager.data["custom_filename_active"] = False

    def select_custom_filename(self):
        self.chk_fn_custom.select()
        self.chk_fn_default.deselect()
        self.config_manager.data["custom_filename_active"] = True

    def get_filename_tokens(self):
        tokens = [
            "Sample Name",
            "Data do Arquivo",
            "Data da Conversão",
            "Hora da Conversão",
            "Timestamp da Conversão",
            "Timestamp ddmmaaaahhmm",
            "Original Files",
            "Layout"
        ]
        for col in self.config_manager.data.get("colunas", []):
            name = col.get("name")
            if name and name not in tokens:
                tokens.append(name)
        for group in ("produtos", "formulas", "extras"):
            for item in self.config_manager.data.get(group, []):
                name = item.get("name")
                if name and name not in tokens:
                    tokens.append(name)
        return tokens

    def on_filename_template_keyrelease(self, event):
        if event.keysym != "bracketleft" and event.char != "[":
            return

        menu = tk.Menu(self, tearoff=0, bg="white", fg=TEXTO, activebackground="#EDF2F7", activeforeground=TEXTO)
        for token in self.get_filename_tokens():
            menu.add_command(label=token, command=lambda t=token: self.insert_filename_token(t))
        try:
            x = self.ent_fn_mask.winfo_rootx() + 10
            y = self.ent_fn_mask.winfo_rooty() + self.ent_fn_mask.winfo_height()
            menu.tk_popup(x, y)
        finally:
            menu.grab_release()

    def insert_filename_token(self, token):
        value = self.ent_fn_mask.get()
        cursor = self.ent_fn_mask.index("insert")
        if cursor > 0 and value[cursor - 1] == "[":
            start = cursor - 1
            new_value = value[:start] + f"[{token}]" + value[cursor:]
            new_cursor = start + len(token) + 2
        else:
            new_value = value[:cursor] + f"[{token}]" + value[cursor:]
            new_cursor = cursor + len(token) + 2
        self.ent_fn_mask.delete(0, "end")
        self.ent_fn_mask.insert(0, new_value)
        self.ent_fn_mask.icursor(new_cursor)

    def create_canva_checkbox(self, parent, text, config_key, check_val=None, invert_checked=False, pack_it=True):
        chk = ctk.CTkCheckBox(parent, text=text, font=FONT_BODY_BOLD,
                              text_color=TEXTO,
                              fg_color=AZUL_NAV_SELECTED,
                              border_color="#90A4BC",
                              checkmark_color="white",
                              hover_color="#1E3E62",
                              bg_color="transparent",
                              corner_radius=5)
        if pack_it:
            chk.pack(anchor="w", padx=25, pady=5)

        # Set checked state
        current = self.config_manager.data.get(config_key, True) if check_val is None else check_val
        if invert_checked:
            current = not current

        if current:
            chk.select()
        else:
            chk.deselect()

        return chk

    def draw_formulas_editor_rows(self):
        for widget in self.formula_rows_container.winfo_children():
            widget.destroy()

        for idx, formula in enumerate(self.config_manager.data["formulas"]):
            row_frame = ctk.CTkFrame(self.formula_rows_container, fg_color="transparent")
            row_frame.pack(fill="x", pady=4)

            # Name entry field — white bg, dark text
            ent_name = ctk.CTkEntry(row_frame, width=120, height=30,
                                    fg_color="white", border_color=BORDA,
                                    text_color=TEXTO, corner_radius=8, font=FONT_BODY)
            ent_name.pack(side="left")
            ent_name.insert(0, formula["name"])

            # Expression entry field — white bg, dark text
            ent_expr = ctk.CTkEntry(row_frame, height=30,
                                    fg_color="white", border_color=BORDA,
                                    text_color=TEXTO, corner_radius=8, font=FONT_BODY)
            ent_expr.pack(side="left", fill="x", expand=True, padx=15)
            ent_expr.insert(0, formula["expression"])

            # Save on focus out
            ent_name.bind("<FocusOut>", lambda e, i=idx, ent=ent_name: self.update_formula_field(i, "name", ent.get()))
            ent_expr.bind("<FocusOut>", lambda e, i=idx, ent=ent_expr: self.update_formula_field(i, "expression", ent.get()))

            # Trash button (orange text outline-style)
            btn_del = ctk.CTkButton(row_frame, text="🗑️", width=25, height=28, fg_color="transparent", 
                                     hover_color="#EDF2F7", text_color="#E06F00", font=("Segoe UI", 12, "bold"),
                                     command=lambda i=idx: self.delete_formula_row(i))
            btn_del.pack(side="right")
            
            # Plus button on the very last row (blue circle with white plus)
            if idx == len(self.config_manager.data["formulas"]) - 1:
                btn_add = ctk.CTkButton(row_frame, text="+", width=28, height=28, 
                                         fg_color="#2E79B9", hover_color="#235A8A", text_color="white", font=("Segoe UI", 12, "bold"),
                                         corner_radius=14, command=self.add_empty_formula_row)
                btn_add.pack(side="right", padx=(5, 0))

        # Handle empty formulas state
        if not self.config_manager.data["formulas"]:
            row_frame = ctk.CTkFrame(self.formula_rows_container, fg_color="transparent")
            row_frame.pack(fill="x", pady=4)
            btn_add = ctk.CTkButton(row_frame, text="Adicionar Fórmula ➕", fg_color="transparent", 
                                     text_color="#2E79B9", font=FONT_BODY_BOLD, command=self.add_empty_formula_row)
            btn_add.pack(anchor="w")

    def update_formula_field(self, idx, key, val):
        if idx < len(self.config_manager.data["formulas"]):
            self.config_manager.data["formulas"][idx][key] = val.strip()

    def add_empty_formula_row(self):
        self.config_manager.data["formulas"].append({
            "name": "Nome",
            "expression": "formula",
            "enabled": True
        })
        self.draw_formulas_editor_rows()

    def delete_formula_row(self, idx):
        if idx < len(self.config_manager.data["formulas"]):
            self.config_manager.data["formulas"].pop(idx)
            self.draw_formulas_editor_rows()

    def save_general_settings(self):
        self.focus() # Trigger focus out of inputs to save texts
        self.config_manager.data["adicionar_respiro"] = bool(self.chk_respiro.get())
        self.config_manager.data["zerar_negativos"] = bool(self.chk_zerar.get())
        self.config_manager.data["data_invertida"] = bool(self.chk_data_inv.get())
        self.config_manager.data["incluir_hora_arquivo"] = bool(self.chk_hora.get())
        self.config_manager.data["custom_filename_active"] = bool(self.chk_fn_custom.get())
        self.config_manager.data["custom_filename_mask"] = self.ent_fn_mask.get().strip()
        self.config_manager.data["separador_campo"] = self.opt_sep_campo.get()
        self.config_manager.data["separador_decimal"] = self.opt_sep_dec.get()
        self.config_manager.data["casas_decimais"] = int(self.opt_dec_places.get())
        
        self.config_manager.save()
        messagebox.showinfo("Sucesso", "Configurações gerais salvas!")

    # --- SCREEN 3: UNIFIED SCROLLABLE COLUMNS LIST ---
    def create_columns_frame(self):
        frame = ctk.CTkFrame(self.content_container, fg_color="transparent")
        
        card = ctk.CTkFrame(frame,
                            fg_color=FUNDO_SUPERFICIE,
                            corner_radius=CARD_RADIUS,
                            border_width=1,
                            border_color=BORDA)
        card.pack(fill="both", expand=True, pady=(0, 15))

        # CARD HEADER - CONFIGURAÇÃO DAS COLUNAS
        card_hdr = ctk.CTkFrame(card, fg_color=CRIMSON_HDR, height=40, corner_radius=CONTROL_RADIUS)
        card_hdr.pack(fill="x", padx=10, pady=10)
        card_hdr.pack_propagate(False)
        ctk.CTkLabel(card_hdr, text="CONFIGURAÇÃO DAS COLUNAS", font=("Segoe UI", 11, "bold"), text_color="white").pack(expand=True)

        self.columns_list_container = ctk.CTkScrollableFrame(
            card,
            fg_color="transparent",
            scrollbar_button_color=BORDA,
            scrollbar_button_hover_color="#90A4BC"
        )
        self.columns_list_container.pack(fill="both", expand=True, padx=5, pady=(0, 5))

        # Save Button bottom right
        btn_save = ctk.CTkButton(frame, text="SALVAR", height=40, width=150,
                                 fg_color=GREEN_SUCCESS, hover_color="#218838",
                                 font=("Segoe UI", 13, "bold"), command=self.save_columns_configuration,
                                 corner_radius=CONTROL_RADIUS)
        btn_save.pack(side="right")

        return frame

    def refresh_columns_lists(self):
        # Clear items
        for widget in self.columns_list_container.winfo_children():
            widget.destroy()
            
        self.row_widgets = {"colunas": [], "produtos": [], "formulas": [], "extras": []}

        # 1. Metadata Section (Default headers)
        for idx, col in enumerate(self.config_manager.data["colunas"]):
            self.draw_column_row(col, "colunas", idx, self.columns_list_container)

        # 2. Products Section Divider
        prod_div = ctk.CTkFrame(self.columns_list_container, fg_color="#D1E0EC", height=28, corner_radius=CONTROL_RADIUS)
        prod_div.pack(fill="x", pady=10)
        prod_div.pack_propagate(False)
        ctk.CTkLabel(prod_div, text="Produtos", font=("Segoe UI Semibold", 10, "bold"), text_color=TEXTO_NAV).pack(expand=True)

        # Products Rows
        for idx, prod in enumerate(self.config_manager.data["produtos"]):
            self.draw_column_row(prod, "produtos", idx, self.columns_list_container)

        # 3. Formulas Section Divider
        form_div = ctk.CTkFrame(self.columns_list_container, fg_color="#D1E0EC", height=28, corner_radius=CONTROL_RADIUS)
        form_div.pack(fill="x", pady=10)
        form_div.pack_propagate(False)
        ctk.CTkLabel(form_div, text="Parâmetros Calculados", font=("Segoe UI Semibold", 10, "bold"), text_color=TEXTO_NAV).pack(expand=True)

        # Formulas Rows
        for idx, formula in enumerate(self.config_manager.data["formulas"]):
            self.draw_column_row(formula, "formulas", idx, self.columns_list_container, has_delete=True)

        # 4. Extras Section Divider
        ext_div = ctk.CTkFrame(self.columns_list_container, fg_color="#D1E0EC", height=28, corner_radius=CONTROL_RADIUS)
        ext_div.pack(fill="x", pady=10)
        ext_div.pack_propagate(False)
        ctk.CTkLabel(ext_div, text="Extras", font=("Segoe UI Semibold", 10, "bold"), text_color=TEXTO_NAV).pack(expand=True)

        # Extras Rows
        for idx, ext in enumerate(self.config_manager.data.get("extras", [])):
            self.draw_column_row(ext, "extras", idx, self.columns_list_container, has_delete=True, is_extra=True)

        # Draw empty extra input row at the very bottom
        self.draw_empty_extra_row(self.columns_list_container)

    def draw_column_row(self, col, list_key, idx, parent, has_delete=False, is_extra=False):
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", pady=4)
        self.row_widgets[list_key].append(row)

        # Move up/down buttons (discreet design)
        btn_up = ctk.CTkButton(row, text="▲", width=20, height=20, fg_color="transparent",
                               hover_color="#EDF2F7", text_color="#A0B0C0", font=("Segoe UI", 10),
                               command=lambda k=list_key, c=col: self.move_list_item(k, c, -1))
        btn_up.pack(side="left", padx=(0, 2))
        
        btn_down = ctk.CTkButton(row, text="▼", width=20, height=20, fg_color="transparent",
                                 hover_color="#EDF2F7", text_color="#A0B0C0", font=("Segoe UI", 10),
                                 command=lambda k=list_key, c=col: self.move_list_item(k, c, 1))
        btn_down.pack(side="left", padx=(0, 5))

        # Custom Checked box
        chk = ctk.CTkCheckBox(row, text="", width=20, fg_color=AZUL_NAV_SELECTED, border_color=AZUL_NAV_SELECTED, corner_radius=6)
        chk.pack(side="left", padx=5)
        if col.get("enabled", True):
            chk.select()
        chk.configure(command=lambda c=col, ch=chk: c.update({"enabled": bool(ch.get())}))

        # Text input/label wide box with light blue border
        box_frame = ctk.CTkFrame(row, fg_color=FUNDO_SUPERFICIE, height=32, corner_radius=CONTROL_RADIUS, border_width=1, border_color=BORDA)
        box_frame.pack(side="left", fill="x", expand=True, padx=5)
        box_frame.pack_propagate(False)
        
        if is_extra:
            # Entry for editing custom extra column title
            ent = ctk.CTkEntry(box_frame, border_width=0,
                               fg_color="white", text_color=TEXTO,
                               font=FONT_BODY_BOLD)
            ent.pack(fill="both", expand=True, padx=15)
            ent.insert(0, col["name"])
            ent.bind("<FocusOut>", lambda e, c=col, ent=ent: c.update({"name": ent.get().strip()}))
        else:
            # Display name label
            lbl = ctk.CTkLabel(box_frame, text=col["name"], font=FONT_BODY_BOLD,
                               text_color=TEXTO, anchor="w", bg_color="transparent")
            lbl.pack(side="left", padx=15)

        # Gear button configuration
        if not is_extra:
            btn_gear = ctk.CTkButton(row, text="⚙️", width=25, height=32, fg_color="transparent", 
                                     hover_color="#D1E0EC", text_color="#2E79B9", font=("Segoe UI", 14, "bold"),
                                     command=lambda c=col, k=list_key: self.open_gear_dialog(c, k))
            btn_gear.pack(side="left", padx=5)

        # Delete button (orange)
        if has_delete:
            btn_del = ctk.CTkButton(row, text="🗑️", width=25, height=32, fg_color="transparent", 
                                     hover_color="#EDF2F7", text_color="#E06F00", font=("Segoe UI", 12, "bold"),
                                     command=lambda k=list_key, c=col: self.delete_list_item(k, c))
            btn_del.pack(side="left", padx=5)

    def draw_empty_extra_row(self, parent):
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", pady=4)

        # Stack placeholder
        ctk.CTkFrame(row, fg_color="transparent", width=25, height=32).pack(side="left")

        # Unchecked checkbox
        chk = ctk.CTkCheckBox(row, text="", width=20, fg_color=AZUL_NAV_SELECTED, border_color=AZUL_NAV_SELECTED, corner_radius=6)
        chk.pack(side="left", padx=5)
        chk.deselect()

        # Empty entry wide frame
        box_frame = ctk.CTkFrame(row, fg_color=FUNDO_SUPERFICIE, height=32, corner_radius=CONTROL_RADIUS, border_width=1, border_color=BORDA)
        box_frame.pack(side="left", fill="x", expand=True, padx=5)
        box_frame.pack_propagate(False)
        
        ent_new = ctk.CTkEntry(box_frame, border_width=0,
                               fg_color="white", text_color=TEXTO_SUAVE,
                               placeholder_text="Escreva para adicionar coluna extra...",
                               font=FONT_BODY)
        ent_new.pack(fill="both", expand=True, padx=15)

        # Trash icon (orange)
        btn_del = ctk.CTkButton(row, text="🗑️", width=25, height=32, fg_color="transparent", 
                                 text_color="#E06F00", text_color_disabled="#E06F00", font=("Segoe UI", 12, "bold"), state="disabled")
        btn_del.pack(side="left", padx=5)

        # Plus button (blue circle with white plus)
        btn_add = ctk.CTkButton(row, text="+", width=28, height=28, fg_color="#2E79B9", 
                                 hover_color="#235A8A", text_color="white", font=("Segoe UI", 12, "bold"),
                                 corner_radius=20,
                                 command=lambda ent=ent_new: self.add_extra_column_row(ent.get()))
        btn_add.pack(side="left", padx=5)

    def add_extra_column_row(self, name):
        name = name.strip()
        if not name:
            messagebox.showwarning("Aviso", "Escreva um nome para a coluna extra!")
            return
            
        if "extras" not in self.config_manager.data:
            self.config_manager.data["extras"] = []
            
        self.config_manager.data["extras"].append({
            "name": name,
            "enabled": True
        })
        self.refresh_columns_lists()

    def open_gear_dialog(self, item, list_key):
        item_type = "product" if list_key == "produtos" else "metadata"
        GearConfigDialog(self, item, item_type, self.config_manager, self.refresh_columns_lists)

    def move_list_item(self, key, col, direction):
        lst = self.config_manager.data[key]
        try:
            index = next(i for i, v in enumerate(lst) if v is col)
        except StopIteration:
            return
            
        new_idx = index + direction
        if 0 <= new_idx < len(lst):
            # Swap data
            lst[index], lst[new_idx] = lst[new_idx], lst[index]
            
            # Repack visual rows instantly
            rows = self.row_widgets.get(key)
            if rows:
                rows[index], rows[new_idx] = rows[new_idx], rows[index]
                if direction == -1:
                    rows[new_idx].pack(before=rows[index])
                else:
                    rows[new_idx].pack(after=rows[index])

    def delete_list_item(self, key, col):
        lst = self.config_manager.data[key]
        try:
            index = next(i for i, v in enumerate(lst) if v is col)
            lst.pop(index)
            self.refresh_columns_lists()
        except StopIteration:
            pass

    def save_columns_configuration(self):
        self.focus()
        self.config_manager.save()
        messagebox.showinfo("Sucesso", "Configuração de colunas e mapeamento salva!")

    # --- SCREEN 4: TABLE PREVIEW ---
    def create_preview_frame(self):
        frame = ctk.CTkFrame(self.content_container, fg_color="transparent")
        
        # Header Crimson Bar
        header = ctk.CTkFrame(frame, fg_color=CRIMSON_HDR, height=40, corner_radius=CONTROL_RADIUS)
        header.pack(fill="x", pady=(0, 12))
        header.pack_propagate(False)
        ctk.CTkLabel(header, text="PRÉVIA DA TABELA", font=("Segoe UI", 16, "bold"), text_color="white").pack(side="left", padx=20, expand=True)
        
        info_frame = ctk.CTkFrame(header, fg_color="transparent")
        info_frame.pack(side="right", padx=20)
        self.lbl_prev_info = ctk.CTkLabel(info_frame, text="Separador: ;  |  Decimal: ,", font=FONT_CAPTION, text_color="white")
        self.lbl_prev_info.pack()

        # Treeview Container Card
        self.table_container = ctk.CTkFrame(frame, fg_color="white", corner_radius=CARD_RADIUS, border_width=1, border_color=BORDA)
        self.table_container.pack(fill="both", expand=True)

        return frame

    def refresh_preview(self):
        sep_campo = self.config_manager.data.get("separador_campo", ";")
        sep_dec = self.config_manager.data.get("separador_decimal", ",")
        try:
            decimal_places = int(self.config_manager.data.get("casas_decimais", 2))
        except (TypeError, ValueError):
            decimal_places = 2
        self.lbl_prev_info.configure(text=f"Separador: {sep_campo}  |  Decimal: {sep_dec}  |  Casas: {decimal_places}")

        # Clear container
        for widget in self.table_container.winfo_children():
            widget.destroy()

        # Resolve Columns in exact unified order
        cols = []
        idx_enabled = False
        idx_name = "ID#"
        for col in self.config_manager.data["colunas"]:
            if col["name"] == "ID#":
                idx_enabled = col.get("enabled", True)
                idx_name = col.get("name", "ID#")
                continue
            if col.get("enabled", True):
                cols.append(col["name"])
        for prod in self.config_manager.data["produtos"]:
            if prod.get("enabled", True):
                cols.append(prod["name"])
        for formula in self.config_manager.data["formulas"]:
            if formula.get("enabled", True):
                cols.append(formula["name"])
        for ext in self.config_manager.data.get("extras", []):
            if ext.get("enabled", True):
                cols.append(ext["name"])

        if idx_enabled:
            cols.insert(0, idx_name)

        if not cols:
            ctk.CTkLabel(self.table_container, text="Nenhuma coluna habilitada!", font=FONT_BODY_BOLD, text_color="#E53E3E").pack(pady=40)
            return

        preview_bar = ctk.CTkFrame(self.table_container, fg_color="#F7FAFC", height=46, corner_radius=CONTROL_RADIUS)
        preview_bar.grid(row=0, column=0, columnspan=2, sticky="ew", padx=14, pady=(14, 6))
        preview_bar.pack_propagate(False)
        ctk.CTkLabel(preview_bar, text="Modelo de saída CSV", font=FONT_BODY_BOLD, text_color=AZUL_NAV_SELECTED).pack(side="left", padx=16)
        ctk.CTkLabel(preview_bar, text=f"{len(cols)} colunas ativas", font=FONT_CAPTION, text_color=TEXTO_SUAVE).pack(side="right", padx=16)

        # Setup Styled Treeview - Canva reference (white headers, dark text, clean zebra)
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview", 
                        background="white", 
                        foreground=TEXTO, 
                        fieldbackground="white", 
                        font=("Segoe UI", 10),
                        rowheight=38,
                        borderwidth=0,
                        highlightthickness=0)
        style.configure("Treeview.Heading", 
                        background="white", 
                        foreground=TEXTO, 
                        font=("Segoe UI Semibold", 9, "bold"),
                        borderwidth=1,
                        relief="flat")
        style.map("Treeview.Heading",
                  background=[('active', "#EDF2F7")])
        style.map("Treeview", 
                  background=[('selected', "#E2E8F0")], 
                  foreground=[('selected', TEXTO)])

        tree = ttk.Treeview(self.table_container, columns=cols, show='headings', selectmode="browse")
        tree.tag_configure('oddrow', background="white")
        tree.tag_configure('evenrow', background="#F4F8FB")
        tree.tag_configure('respirorow', background="#EEF4FA")

        for col in cols:
            tree.heading(col, text=col)
            tree.column(col, width=max(110, min(170, len(col) * 9)), anchor="center")

        # Mock values
        def get_mock_val(cname):
            c = cname.upper()
            if "NOME" in c or "SAMPLE" in c: return "Mosto"
            if "DATA" in c or "DATE" in c: return datetime.datetime.now().strftime("%Y/%m/%d" if self.config_manager.data.get("data_invertida", False) else "%d/%m/%Y")
            if "DIA" in c: return datetime.datetime.now().strftime("%d")
            if "MÊS" in c or "MES" in c or "MONTH" in c: return datetime.datetime.now().strftime("%m")
            if "HORA" in c or "HOUR" in c: return "08:00"
            if "USUÁRIO" in c or "USER" in c: return "Operador_Thermo"
            if "MÉTODO" in c or "METHOD" in c: return "Guarani_2026"
            if "ORIGINAL" in c: return "24,0000 (1).xls"
            
            # Numeric defaults
            v_map = {"SACAROSE": 10.75, "GLICOSE": 1.41, "FRUTOSE": 2.04, "MANITOL": 0.03, "GLICEROL": 0.08, "ETANOL": 0.00, "AR": 3.45, "ART": 14.78}
            base_val = v_map.get(c, 0.0)
            return f"{base_val:.{decimal_places}f}".replace(".", sep_dec)

        # Build list representing preview rows
        respiro = self.config_manager.data.get("adicionar_respiro", True)
        if respiro:
            tree.insert("", "end", values=[""] * len(cols), tags=('respirorow',))

        # Mock sample row
        sample_row = []
        row_id = 1
        for col in cols:
            if col == idx_name:
                sample_row.append(str(row_id))
            else:
                sample_row.append(get_mock_val(col))
        tree.insert("", "end", values=sample_row, tags=('oddrow',))

        # Scrollbars
        vsb = ctk.CTkScrollbar(self.table_container, orientation="vertical", command=tree.yview, width=12)
        hsb = ctk.CTkScrollbar(self.table_container, orientation="horizontal", command=tree.xview, height=12)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        tree.grid(row=1, column=0, sticky="nsew", padx=(14, 0), pady=(6, 0))
        vsb.grid(row=1, column=1, sticky="ns", padx=(8, 14), pady=(6, 0))
        hsb.grid(row=2, column=0, sticky="ew", padx=(14, 0), pady=(8, 14))
        
        self.table_container.grid_columnconfigure(0, weight=1)
        self.table_container.grid_rowconfigure(1, weight=1)

    # --- ACTION CONTROLS ---
    def refresh_general_configs(self):
        def set_chk(chk, val):
            chk.select() if val else chk.deselect()
            
        set_chk(self.chk_respiro, self.config_manager.data.get("adicionar_respiro", False))
        set_chk(self.chk_zerar, self.config_manager.data.get("zerar_negativos", True))
        set_chk(self.chk_data_inv, self.config_manager.data.get("data_invertida", False))
        set_chk(self.chk_hora, self.config_manager.data.get("incluir_hora_arquivo", True))
        
        custom_fn = self.config_manager.data.get("custom_filename_active", False)
        set_chk(self.chk_fn_custom, custom_fn)
        set_chk(self.chk_fn_default, not custom_fn)
        if custom_fn:
            self.select_custom_filename()
        else:
            self.select_default_filename()
            
        self.ent_fn_mask.delete(0, "end")
        self.ent_fn_mask.insert(0, self.config_manager.data.get("custom_filename_mask", "[Sample Name]_[Timestamp da Conversão]"))
        
        self.opt_sep_campo.set(self.config_manager.data.get("separador_campo", ";"))
        self.opt_sep_dec.set(self.config_manager.data.get("separador_decimal", ","))
        self.opt_dec_places.set(str(self.config_manager.data.get("casas_decimais", 4)))
        
        self.draw_formulas_editor_rows()

    def reset_defaults(self):
        if messagebox.askyesno("Confirmar", "Deseja restaurar as configurações de colunas, produtos e fórmulas originais?"):
            self.config_manager.reset()
            self.refresh_columns_lists()
            self.refresh_general_configs()
            self.refresh_preview()
            self.update_folder_counts()
            messagebox.showinfo("Sucesso", "Configurações restauradas para o padrão.")

    def show_help(self):
        help_win = ctk.CTkToplevel(self)
        help_win.title("CETEC | Central de Ajuda")
        help_win.geometry("620x600")
        help_win.configure(fg_color=FUNDO_APP)
        help_win.grab_set()

        self.after(200, lambda: self.apply_help_icon(help_win))

        header = ctk.CTkFrame(help_win, height=60, fg_color=CRIMSON_HDR, corner_radius=0)
        header.pack(fill="x")
        ctk.CTkLabel(header, text="Modo de Uso — Conversor Thermo", font=("Segoe UI", 14, "bold"), text_color="white").pack(side="left", padx=20)

        container = ctk.CTkScrollableFrame(help_win, fg_color=FUNDO_SUPERFICIE, corner_radius=14, border_width=1, border_color=BORDA)
        container.pack(fill="both", expand=True, padx=20, pady=(20, 10))

        FONT_SEC = ("Segoe UI", 12, "bold")
        FONT_TXT = ("Segoe UI", 11)
        W = 560

        sections = [
            (
                "1. Configurar as pastas",
                "Vá até a aba Inicio e informe a Pasta de Origem (onde chegam os arquivos XLS do cromatografo) "
                "e a Pasta de Destino (onde sera salvo o CSV final). "
                "A Pasta de Backup e opcional: se preenchida, o arquivo original e copiado para la apos a conversao."
            ),
            (
                "2. Converter manualmente",
                "Clique no botao Manual (laranja). O programa processa todos os arquivos XLS presentes na pasta "
                "de origem de uma vez e salva os CSVs na pasta de destino."
            ),
            (
                "3. Modo Sentinela",
                "Clique em Sentinela (azul) para ativar o monitoramento continuo. O programa fica em espera e "
                "converte automaticamente qualquer novo arquivo que aparecer na pasta de origem. "
                "Para encerrar o monitoramento, clique no botao novamente."
            ),
            (
                "4. Configuracoes Gerais",
                "Na aba Configuracoes Gerais voce pode:\n"
                "  - Ativar a linha de respiro (linha vazia entre amostras para sincronizar com o Autolab)\n"
                "  - Zerar valores negativos\n"
                "  - Escolher o separador de campo e decimal\n"
                "  - Definir quantas casas decimais usar\n"
                "  - Personalizar o nome do arquivo de saida"
            ),
            (
                "5. Colunas e Produtos",
                "Na aba Configuracoes de Colunas voce pode renomear os picos (ex.: Glucose para Glicose), "
                "reordenar as colunas com os botoes de seta e ativar ou desativar produtos e formulas."
            ),
            (
                "6. Restaurar Padroes",
                "O botao Restaurar Padroes devolve todas as configuracoes ao estado original definido pelo administrador. "
                "Use com cuidado, pois desfaz todas as alteracoes feitas."
            ),
            (
                "Suporte",
                "Em caso de duvidas ou problemas, entre em contato com o suporte tecnico responsavel."
            ),
        ]

        for title, body in sections:
            ctk.CTkLabel(container, text=title, font=FONT_SEC, text_color=CRIMSON_HDR, anchor="w", wraplength=W).pack(anchor="w", padx=20, pady=(14, 2))
            ctk.CTkLabel(container, text=body, font=FONT_TXT, text_color=TEXTO, anchor="w", justify="left", wraplength=W).pack(anchor="w", padx=20, pady=(0, 4))

        ctk.CTkButton(help_win, text="ENTENDI", command=help_win.destroy, fg_color=GREEN_SUCCESS, hover_color="#218838", height=40, corner_radius=8).pack(pady=(0, 20))

    def apply_help_icon(self, win):
        icon_path = resource_path("logo.ico")
        if os.path.exists(icon_path):
            try: win.iconbitmap(icon_path)
            except: pass

    # --- CONVERSION ACTIONS ---
    def run_manual_conversion(self):
        if self.sentinel_active:
            messagebox.showwarning("Aviso", "Desligue o monitoramento antes de converter manualmente.")
            return

        orig = self.config_manager.data.get("pasta_origem", "")
        dest = self.config_manager.data.get("pasta_destino", "")
        if not orig or not dest or not os.path.exists(orig) or not os.path.exists(dest):
            messagebox.showwarning("Aviso", "Configure as pastas de origem e destino no painel principal!")
            return
            
        self.btn_manual_convert.configure(state="disabled")
        self.btn_toggle_sentinel.configure(state="disabled")
        self.progress_lbl.configure(text="Iniciando conversão manual...")
        self.progress_bar.set(0.1)
        
        threading.Thread(target=self.process_manual_task, daemon=True).start()

    def process_manual_task(self):
        try:
            orig = self.config_manager.data.get("pasta_origem", "")
            dest = self.config_manager.data.get("pasta_destino", "")
            back = self.config_manager.data.get("pasta_backup", "")
            
            files = [f for f in os.listdir(orig) if f.lower().endswith(('.xls', '.xlsx'))]
            if not files:
                self.progress_lbl.configure(text="Nenhum arquivo XLS encontrado na pasta de origem.")
                self.progress_bar.set(0)
                self.btn_manual_convert.configure(state="normal")
                return

            total = len(files)
            success_count = 0
            
            for idx, file in enumerate(files):
                self.progress_lbl.configure(text=f"Processando {file}...")
                self.progress_bar.set(float(idx / total))
                self.update_idletasks()

                src_path = os.path.join(orig, file)
                success = self.parser.convert_file(src_path, dest)
                if success:
                    success_count += 1
                    if back:
                        if not os.path.exists(back):
                            os.makedirs(back)
                        shutil.copy(src_path, os.path.join(back, file))
                    try:
                        os.remove(src_path)
                    except Exception as e:
                        print(f"Error removing source: {e}")

            self.progress_lbl.configure(text=f"Sucesso! {success_count} de {total} arquivos convertidos.")
            self.progress_bar.set(1.0)
            self.update_folder_counts()
            messagebox.showinfo("Conversão Concluída", f"{success_count} arquivos processados com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro: {e}")
        finally:
            self.btn_manual_convert.configure(state="normal")
            if not self.sentinel_active:
                self.btn_toggle_sentinel.configure(state="normal")

    # --- SENTINELA WATCHDOG loop ---
    def toggle_sentinel(self):
        if self.sentinel_active:
            # STOP WATCHDOG
            self.sentinel_active = False
            self.lbl_sentinel_indicator.pack_forget()
            self.btn_toggle_sentinel.configure(text="▶", fg_color=BLUE_WATCH_BTN, hover_color="#638CBA")
            self.btn_manual_convert.configure(state="normal")
        else:
            # START WATCHDOG
            orig = self.config_manager.data.get("pasta_origem", "")
            dest = self.config_manager.data.get("pasta_destino", "")
            if not orig or not dest or not os.path.exists(orig) or not os.path.exists(dest):
                messagebox.showwarning("Aviso", "Configure as pastas de origem e destino no painel principal antes de ligar!")
                return

            self.sentinel_active = True
            self.converted_count = 0
            self.lbl_sentinel_indicator.configure(text="Ativo | 0 convertidos")
            self.lbl_sentinel_indicator.pack(pady=(5, 0))
            self.btn_toggle_sentinel.configure(text="||", fg_color=GREEN_SUCCESS, hover_color="#218838")
            self.btn_manual_convert.configure(state="disabled")
            
            # Spawn loop
            self.sentinel_thread = threading.Thread(target=self.sentinel_monitor_loop, daemon=True)
            self.sentinel_thread.start()

    def sentinel_monitor_loop(self):
        while self.sentinel_active:
            try:
                orig = self.config_manager.data.get("pasta_origem", "")
                dest = self.config_manager.data.get("pasta_destino", "")
                back = self.config_manager.data.get("pasta_backup", "")

                if not orig or not dest or not os.path.exists(orig) or not os.path.exists(dest):
                    self.after(0, self.toggle_sentinel)
                    break

                files = [f for f in os.listdir(orig) if f.lower().endswith(('.xls', '.xlsx'))]
                
                for file in files:
                    src_path = os.path.join(orig, file)
                    try:
                        initial_size = os.path.getsize(src_path)
                        time.sleep(0.5)
                        if os.path.getsize(src_path) != initial_size:
                            continue
                    except Exception:
                        continue

                    # Process file
                    success = self.parser.convert_file(src_path, dest)
                    if success:
                        self.converted_count += 1
                        self.after(0, self.update_folder_counts)
                        
                        if back:
                            if not os.path.exists(back):
                                os.makedirs(back)
                            shutil.copy(src_path, os.path.join(back, file))
                        
                        try:
                            os.remove(src_path)
                        except Exception as e:
                            print(f"Error removing sentinel file: {e}")
            except Exception as e:
                print(f"Sentinel Loop Error: {e}")
                
            time.sleep(2)

if __name__ == "__main__":
    app = ThermoConverterApp()
    app.mainloop()
