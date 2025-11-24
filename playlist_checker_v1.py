import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from rapidfuzz import fuzz
from unidecode import unidecode
import re
from datetime import datetime
import warnings
import shutil
import json

# Suppress openpyxl warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# Constants
MIN_MATCH_SCORE = 70
HET_TAP_MIN_MATCH_SCORE = 100

# Replacements for ten_CT hdps 
REPLACEMENTS = {}

# Ad paths configuration
CHANNEL_3_AD_PATHS = [
    "\\\\server-40t02\\thanhpham$\\P. Quang cao\\THVL3\\CMSP_VL3\\CMSP_THVL3_%d.%m.%Y",
    "\\\\server-40t02\\thanhpham$\\P. Quang cao\\THVL3\\Quang cao VL3\\QC-THVL3__%d-%m-%Y",
    "\\\\server-40t02\\thanhpham$\\P. Quang cao\\THVL3\\ThongBaoVL3\\%d-%m-%Y"
]

CHANNEL_4_AD_PATHS = [
    "\\\\server-40t02\\thanhpham$\\P. Quang cao\\THVL4\\CMSP_VL4\\CMSP_THVL4_%d.%m.%Y",
    "\\\\server-40t02\\thanhpham$\\P. Quang cao\\THVL4\\ThongBao VL4\\%d-%m-%Y"
]

class CheckPlaylistApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Check Playlist")
        self.root.geometry("600x500")
        
        self.excel_path_hdps = tk.StringVar()
        self.excel_path_playlist = tk.StringVar()
        self.selected_titles = tk.StringVar(value="sang")
        self.channel = None
        self.date = None
        
        # Frame cho input files
        input_frame = tk.Frame(root)
        input_frame.pack(padx=10, pady=5, fill="x")
        
        tk.Label(input_frame, text="File Excel HDPS:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        tk.Entry(input_frame, textvariable=self.excel_path_hdps, width=40).grid(row=0, column=1, padx=5, pady=5)
        tk.Button(input_frame, text="Browse", command=self.browse_excel_hdps).grid(row=0, column=2, padx=5, pady=5)
        
        tk.Label(input_frame, text="File Excel Playlist:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        tk.Entry(input_frame, textvariable=self.excel_path_playlist, width=40).grid(row=1, column=1, padx=5, pady=5)
        tk.Button(input_frame, text="Browse", command=self.browse_excel_playlist).grid(row=1, column=2, padx=5, pady=5)
        
        # Frame cho radio buttons
        radio_frame = tk.Frame(root)
        radio_frame.pack(padx=10, pady=5)
        
        tk.Label(radio_frame, text="Chon playlist:").pack(side="left", padx=5)
        tk.Radiobutton(radio_frame, text="Sang", variable=self.selected_titles, value="sang").pack(side="left", padx=5)
        tk.Radiobutton(radio_frame, text="Trua", variable=self.selected_titles, value="trua").pack(side="left", padx=5)
        tk.Radiobutton(radio_frame, text="Chieu", variable=self.selected_titles, value="chieu").pack(side="left", padx=5)
        tk.Radiobutton(radio_frame, text="All", variable=self.selected_titles, value="all").pack(side="left", padx=5)
        
        # Button check
        tk.Button(root, text="Check Playlist", command=self.check_playlist, bg="#4CAF50", fg="white", font=("Arial", 10, "bold")).pack(pady=10)
        
        # Text area để hiển thị kết quả
        tk.Label(root, text="Ket qua:", font=("Arial", 10, "bold")).pack(padx=10, pady=(10, 0), anchor="w")
        self.result_text = scrolledtext.ScrolledText(root, width=70, height=15, wrap=tk.WORD)
        self.result_text.pack(padx=10, pady=5, fill="both", expand=True)
        self.result_text.config(state='disabled')

    def browse_excel_hdps(self):
        file_path = filedialog.askopenfilename(
            title="Chon file Excel HDPS",
            initialdir="C:\\Users\\LISTBOX01\\Downloads",
            filetypes=[("Excel files", "*.xls *.xlsx"), ("All files", "*.*")]
        )
        if file_path:
            self.excel_path_hdps.set(file_path)
            channel_match = re.search(r"THVL(\d+).*", file_path)
            self.channel = channel_match.group(1) if channel_match else None
            
            if not self.channel or self.channel not in ["3", "4"]:
                messagebox.showerror("Loi", "Chi chon HDPS Kenh 3 hoac Kenh 4!")
                self.excel_path_hdps.set("")
                return
                
            date_match = re.search(r"[\s_](\d+).*?\.xls", file_path)
            if date_match:
                self.date = date_match.group(1)
                try:
                    datetime.strptime(self.date, "%d%m%Y")
                except ValueError:
                    messagebox.showerror("Loi", "Dinh dang ngay khong hop le! Dam bao ngay o dinh dang ddmmyyyy.")
                    self.excel_path_hdps.set("")
                    self.date = None
                    return
            else:
                messagebox.showerror("Loi", "Khong tim thay dinh dang ngay trong ten file Excel!")
                self.excel_path_hdps.set("")
                return

    def browse_excel_playlist(self):
        file_path = filedialog.askopenfilename(
            title="Chon file Excel playlist",
            initialdir="D:\\DAN\\test\\playlist_checker_idea",
            filetypes=[("Excel files", "*.xls *.xlsx"), ("All files", "*.*")]
        )
        if file_path:
            self.excel_path_playlist.set(file_path)

    def check_playlist(self):
        excel_path_hdps = self.excel_path_hdps.get()
        excel_path_playlist = self.excel_path_playlist.get()
        selected_titles = self.selected_titles.get()

        if not excel_path_hdps:
            messagebox.showerror("Loi", "Ban chua chon file Excel HDPS!")
            return
        if not excel_path_playlist:
            messagebox.showerror("Loi", "Ban chua chon file Excel playlist!")
            return
        if not self.date:
            messagebox.showerror("Loi", "Ngay khong hop le! Vui long kiem tra lai file Excel.")
            return
        if not os.path.isfile(excel_path_hdps):
            messagebox.showerror("Loi", "File Excel HDPS khong ton tai!")
            return
        if not os.path.isfile(excel_path_playlist):
            messagebox.showerror("Loi", "File Excel playlist khong ton tai!")
            return

        try:
            result = check_hdps_and_playlist(excel_path_hdps, excel_path_playlist, self.date, selected_titles, self.channel)
            self.result_text.config(state='normal')  # Bật lại để có thể cập nhật
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(tk.END, result)
            self.result_text.config(state='disabled')  # Tắt lại sau khi cập nhật
        except Exception as e:
            messagebox.showerror("Loi", f"Loi trong qua trinh check playlist:\n{str(e)}")


def parse_time(time_str):
    """Chuyen doi thoi gian tu string sang so giay"""
    if pd.isna(time_str) or time_str == "":
        return 0
    
    time_str = str(time_str).strip()
    
    # Loai bo prefix ngay moi: (1), (2), (3), ... truoc khi xu ly
    # Format: (1)HH:MM:SS:FF hoac (2)HH:MM:SS:FF
    import re
    time_str = re.sub(r'^\(\d+\)', '', time_str)
    
    # Neu la datetime object tu pandas, chi lay time part
    if isinstance(time_str, str) and 'T' in time_str:
        time_str = time_str.split('T')[1] if 'T' in time_str else time_str
    
    # Loai bo date part neu co (format: "1900-01-01 HH:MM:SS")
    if ' ' in time_str and '-' in time_str:
        parts = time_str.split(' ')
        if len(parts) == 2:
            time_str = parts[1]
    
    # Format HH:MM:SS:FF (gio:phut:giay:frame)
    if time_str.count(':') == 3:
        try:
            parts = time_str.split(':')
            hours = int(parts[0])
            minutes = int(parts[1])
            seconds = int(parts[2])
            frames = int(parts[3])
            return hours * 3600 + minutes * 60 + seconds + frames / 25.0
        except (ValueError, IndexError) as e:
            print(f"Loi parse time (HH:MM:SS:FF): {time_str}, Error: {e}")
            return 0
    
    # Format HH:MM:SS
    elif time_str.count(':') == 2:
        try:
            parts = time_str.split(':')
            hours = int(parts[0])
            minutes = int(parts[1])
            seconds = float(parts[2])
            return hours * 3600 + minutes * 60 + seconds
        except (ValueError, IndexError) as e:
            print(f"Loi parse time (HH:MM:SS): {time_str}, Error: {e}")
            return 0
    
    return 0


def format_time(seconds):
    """Chuyen doi so giay thanh format HH:MM:SS"""
    hours = int(seconds // 3600)
    minutes = int((seconds % 3600) // 60)
    secs = int(seconds % 60)
    return f"{hours:02d}:{minutes:02d}:{secs:02d}"


def normalize_title(title):
    """Chuan hoa ten chuong trinh de so sanh - khong dau lowercase"""
    if pd.isna(title) or title is None or str(title).strip() == '':
        return ""
    
    try:
        title = str(title).strip()
        # Chuyen ve khong dau va lowercase
        title = unidecode(title).lower()
        # Loai bo cac ky tu dac biet, giu lai space
        #title = re.sub(r'[\-]+', ' ', title)
        # Loai bo khoang trang thua
        title = re.sub(r'\s+', ' ', title).strip()
        return title
    except Exception as e:
        print(f"Loi khi chuan hoa title: {title}, Error: {e}")
        return ""


def normalize_hdps_dataframe(df_hdps):
    """Chuan hoa DataFrame HDPS thanh format chuan: [Gio, Thoi luong, Ten chuong trinh, is_new_row]"""
    df_hdps = df_hdps.dropna(how='all').reset_index(drop=True)
    
    df_normalized = pd.DataFrame({
        'Gio': df_hdps.iloc[:, 0],
        'Thoi_luong': df_hdps.iloc[:, 1],
        'Ten_CT': df_hdps.iloc[:, 2]
    })
    
    df_normalized = df_normalized[df_normalized['Ten_CT'].notna()].reset_index(drop=True)
    
    def is_trailer(title):
        if pd.isna(title):
            return False
        title_str = str(title).strip().lower()
        title_str = unidecode(title_str)
        if ':' in title_str:
            title_str = title_str.split(':', 1)[1].strip()
        return title_str.startswith('trailer') or title_str.startswith('qc') or title_str.startswith('thong bao') or title_str.startswith('=>') or "nguoi len lich" in title_str
    
    df_normalized = df_normalized[~df_normalized['Ten_CT'].apply(is_trailer)].reset_index(drop=True)
    
    new_rows = []
    
    for idx, row in df_normalized.iterrows():
        ten_ct = str(row['Ten_CT']).strip()
        gio = row['Gio']
        thoi_luong = row['Thoi_luong']
        
        if ':' in ten_ct:
            parts = ten_ct.split(':', 1)
            prefix = parts[0].strip()
            prefix = unidecode(prefix.lower())
            ten_chinh = parts[1].strip()
            ten_chinh = unidecode(ten_chinh.lower())
            
            if prefix == "phim vn":
                tap_match = re.search(r'(?:[-_\s]*)(?:t|tap|hdgc)\.?[-_\s]*(\d+)(?:_?\w*[-_\s]+(?:t|tap)?\.?[-_\s]*(\d+))?', ten_chinh)
                
                if tap_match:
                    tap_start = int(tap_match.group(1))
                    # Neu khong co group 2 thi tap_end = tap_start
                    tap_end = int(tap_match.group(2)) if tap_match.group(2) else tap_start
                    
                    # Lay ten phim (phan truoc doan match so tap)
                    ten_phim = ten_chinh[:tap_match.start()].strip()
                    # Xoa cac ky tu thua o cuoi ten phim (gach noi, gach duoi)
                    ten_phim = re.sub(r'[-_]+$', '', ten_phim).strip()
                    
                    if tap_start == tap_end:
                        # Truong hop 1 tap don le
                        new_rows.append({'Gio': gio, 'Thoi_luong': thoi_luong, 'Ten_CT': f'Dau phim_{ten_chinh}', 'is_new_row': True, 'is_dau_phim': True})
                        new_rows.append({'Gio': gio, 'Thoi_luong': thoi_luong, 'Ten_CT': ten_chinh, 'is_new_row': True, 'is_dau_phim': False})
                        if "(het)" in ten_chinh:
                            new_rows.append({'Gio': gio, 'Thoi_luong': thoi_luong, 'Ten_CT': f'TAP CUOI', 'is_new_row': True, 'is_dau_phim': False})
                        else:                                
                            new_rows.append({'Gio': gio, 'Thoi_luong': thoi_luong, 'Ten_CT': f'HET TAP {tap_start}', 'is_new_row': True, 'is_dau_phim': False})
                        new_rows.append({'Gio': gio, 'Thoi_luong': thoi_luong, 'Ten_CT': f'Duoi phim_{ten_chinh}', 'is_new_row': True, 'is_dau_phim': False})
                    else:
                        # Truong hop dai tap (VD: 19-20)
                        for tap in range(int(tap_start), int(tap_end) + 1):
                            # Tao ten tap ao cho cac tap trong khoang
                            ten_tap = f'{ten_phim} - t.{tap}'
                            
                            if tap == int(tap_end):
                                new_rows.append({'Gio': gio, 'Thoi_luong': thoi_luong, 'Ten_CT': f'Dau phim_{ten_tap}', 'is_new_row': True, 'is_dau_phim': False})
                            else:
                                new_rows.append({'Gio': gio, 'Thoi_luong': thoi_luong, 'Ten_CT': f'Dau phim_{ten_chinh}', 'is_new_row': True, 'is_dau_phim': True})

                            new_rows.append({'Gio': gio, 'Thoi_luong': thoi_luong, 'Ten_CT': ten_tap, 'is_new_row': True, 'is_dau_phim': False})
                            
                            if "(het)" in ten_chinh:
                                new_rows.append({'Gio': gio, 'Thoi_luong': thoi_luong, 'Ten_CT': f'TAP CUOI', 'is_new_row': True, 'is_dau_phim': False})
                            else:                                
                                new_rows.append({'Gio': gio, 'Thoi_luong': thoi_luong, 'Ten_CT': f'HET TAP {tap}', 'is_new_row': True, 'is_dau_phim': False})
                                
                            new_rows.append({'Gio': gio, 'Thoi_luong': thoi_luong, 'Ten_CT': f'Duoi phim_{ten_tap}', 'is_new_row': True, 'is_dau_phim': False})
                else:
                    # Truong hop khong tim thay so tap (VD: Bi mat cua luat su_hdgc_004)
                    # Giu nguyen dong, khong tach Dau/Duoi
                    new_rows.append({'Gio': gio, 'Thoi_luong': thoi_luong, 'Ten_CT': ten_chinh, 'is_new_row': False, 'is_dau_phim': False})
            
            elif prefix == "phim sitcom" and "trong nha ngoai ngo" not in ten_ct.lower():
                new_rows.append({'Gio': gio, 'Thoi_luong': thoi_luong, 'Ten_CT': 'HH_PHIM TRUYEN VIET NAM', 'is_new_row': True, 'is_dau_phim': True})
                new_rows.append({'Gio': gio, 'Thoi_luong': thoi_luong, 'Ten_CT': f'{ten_chinh}', 'is_new_row': True, 'is_dau_phim': False})
            
            elif prefix == "cai luong":
                new_rows.append({'Gio': gio, 'Thoi_luong': thoi_luong, 'Ten_CT': 'Hinh hieu_San Khau Cai Luong_HD', 'is_new_row': True, 'is_dau_phim': True})
                new_rows.append({'Gio': gio, 'Thoi_luong': thoi_luong, 'Ten_CT': f'{ten_chinh}', 'is_new_row': True, 'is_dau_phim': False})
                if "p.1" in ten_ct.lower():
                    new_rows.append({'Gio': gio, 'Thoi_luong': thoi_luong, 'Ten_CT': f'HET PHAN 1', 'is_new_row': True, 'is_dau_phim': False})
                else:
                    new_rows.append({'Gio': gio, 'Thoi_luong': thoi_luong, 'Ten_CT': f'TAP CUOI', 'is_new_row': True, 'is_dau_phim': False})
            
            elif "hai chang hao hon" in ten_ct.lower() or "gai khon duoc chong" in ten_ct.lower():
                new_rows.append({'Gio': gio, 'Thoi_luong': thoi_luong, 'Ten_CT': 'HINH HIEU CO TICH VN', 'is_new_row': True, 'is_dau_phim': True})
                new_rows.append({'Gio': gio, 'Thoi_luong': thoi_luong, 'Ten_CT': f'{ten_chinh}', 'is_new_row': True, 'is_dau_phim': False}) 
            
            else:
                new_rows.append({'Gio': gio, 'Thoi_luong': thoi_luong, 'Ten_CT': ten_chinh, 'is_new_row': False, 'is_dau_phim': False})
        else:
            new_rows.append({'Gio': gio, 'Thoi_luong': thoi_luong, 'Ten_CT': ten_ct, 'is_new_row': False, 'is_dau_phim': False})
    
    df_result = pd.DataFrame(new_rows)
    
    # Ham apply replacement - SUA LAI DE NORMALIZE TRUOC KHI SO SANH
    def apply_replacements(ten_CT):
        """Ap dung cac replacement rules"""
        if pd.isna(ten_CT):
            return ten_CT
        
        ten_CT = str(ten_CT).strip()
        
        # NORMALIZE ten_CT de so sanh (khong dau, lowercase)
        ten_CT_normalized = normalize_title(ten_CT)
        
        # Duyet qua cac replacement rules
        for original, replacement in REPLACEMENTS.items():
            # NORMALIZE original key de so sanh
            original_normalized = normalize_title(original)
            
            if replacement["replace_type"] == "full":
                # So sanh voi chuoi da normalize
                if original_normalized == ten_CT_normalized:
                    return replacement["value"]
                # Hoac kiem tra xem original co nam trong ten_CT khong
                elif original_normalized in ten_CT_normalized:
                    return replacement["value"]
            elif replacement["replace_type"] == "partial":
                # Thay the phan text (khong phan biet hoa thuong)
                if original_normalized in ten_CT_normalized:
                    # Tim vi tri trong chuoi goc va thay the
                    ten_CT_normalized = ten_CT_normalized.replace(original_normalized, replacement["value"])
        
        return ten_CT_normalized
    
    # Apply replacements cho tat ca cac dong
    df_result['Ten_CT'] = df_result['Ten_CT'].apply(apply_replacements)
    
    return df_result


def normalize_playlist_dataframe(df_playlist):
    """Chuan hoa DataFrame Playlist thanh format chuan: [Gio, Thoi luong, Ten chuong trinh]"""
    # Loai bo cac dong hoan toan trong
    df_playlist = df_playlist.dropna(how='all').reset_index(drop=True)

    # Su dung index
    time_col = df_playlist.columns[1]
    duration_col = df_playlist.columns[2]
    title_col = df_playlist.columns[5]
    
    df_normalized = pd.DataFrame({
        'Gio': df_playlist[time_col],
        'Thoi_luong': df_playlist[duration_col],
        'Ten_CT': df_playlist[title_col]
    })
    
    # Loai bo cac dong co gia tri NaN o cot Ten_CT
    df_normalized = df_normalized[df_normalized['Ten_CT'].notna()].reset_index(drop=True)
    
    # Loc bo cac dong co Type la Note, Event hoac Trailer
    def is_note_or_event_or_trailer_only(title):
        if pd.isna(title):
            return False
        title_str = str(title).strip().lower()
        return '[note]' in title_str or '[event]' in title_str or '[stop]' in title_str or '[gap]' in title_str or 'show logo' in title_str or 'logo off' in title_str or 'hinh hieu ca nhac' in title_str or title_str.startswith('-') or title_str.startswith('trailer')

    df_for_ad_check  = df_normalized[~df_normalized['Ten_CT'].apply(is_note_or_event_or_trailer_only)].reset_index(drop=True)
    
    
    # Loc bo cac dong co qc,tb,cmsp
    def is_qc_or_cmsp_or_tb(title):
        if pd.isna(title):
            return False
        title_str = str(title).strip().lower()
        return title_str.startswith('cmsp_') or title_str.startswith('qc-') or title_str.startswith('tb-')

    df_normalized = df_for_ad_check[~df_for_ad_check['Ten_CT'].apply(is_qc_or_cmsp_or_tb)].reset_index(drop=True)
    

    
    # ==== BUOC GOP CAC PART LIEN KE VA KIEM TRA THU TU ====
    def extract_part_info(title):
        """Trich xuat thong tin part tu title"""
        if pd.isna(title):
            return None, None
        
        title_str = str(title).strip()
        # Tim pattern (Part X) o cuoi title
        match = re.search(r'\(Part\s+(\d+)\)\s*$', title_str, re.IGNORECASE)
        if match:
            part_num = int(match.group(1))
            base_title = title_str[:match.start()].strip()
            return base_title, part_num
        return None, None

    # Danh sach luu cac loi thu tu part
    part_order_errors = []

    # Duyet qua cac dong va danh dau cac dong can giu lai
    rows_to_keep = []
    i = 0

    while i < len(df_normalized):
        current_row = df_normalized.iloc[i]
        current_title = current_row['Ten_CT']
        base_title, part_num = extract_part_info(current_title)
        
        # Neu khong phai la part, giu lai dong nay
        if base_title is None:
            rows_to_keep.append(i)
            i += 1
            continue
        
        # Neu la part, kiem tra xem co the gop voi cac part ke tiep khong
        consecutive_parts = [i]  # Luu chi so cac dong co part lien tiep
        part_numbers = [part_num]  # Luu cac so part
        expected_next_part = part_num + 1
        j = i + 1
        
        # Tim cac part lien tiep cung base_title
        while j < len(df_normalized):
            next_row = df_normalized.iloc[j]
            next_title = next_row['Ten_CT']
            next_base_title, next_part_num = extract_part_info(next_title)
            
            # Kiem tra xem co phai la cung base_title khong
            if next_base_title == base_title:
                consecutive_parts.append(j)
                part_numbers.append(next_part_num)
                
                # Kiem tra thu tu part
                if next_part_num != expected_next_part:
                    part_order_errors.append({
                        'title': base_title,
                        'expected': expected_next_part,
                        'actual': next_part_num,
                        'row': j + 1
                    })
                
                expected_next_part = next_part_num + 1
                j += 1
            else:
                # Gap dong khong phai la cung base_title, dung lai
                break
        
        # Neu chi co 1 part hoac cac part khong lien tiep, giu nguyen
        if len(consecutive_parts) == 1:
            rows_to_keep.append(i)
            i += 1
        else:
            # Chi giu lai dong dau tien (gop tat ca cac part)
            rows_to_keep.append(consecutive_parts[0])
            i = j  # Nhay den dong sau cum part da xu ly

    # Tao DataFrame moi chi chua cac dong can giu lai
    df_result = df_normalized.iloc[rows_to_keep].reset_index(drop=True)

    # Loai bo phan (Part X) khoi cac title da duoc giu lai
    def remove_part_suffix(title):
        if pd.isna(title):
            return title
        title_str = str(title).strip()
        # Loai bo (Part X) o cuoi
        return re.sub(r'\(Part\s+\d+\)\s*$', '', title_str, flags=re.IGNORECASE).strip()

    df_result['Ten_CT'] = df_result['Ten_CT'].apply(remove_part_suffix)

    return df_result, df_for_ad_check, part_order_errors


def get_section_data(df_normalized, section):
    """Lay du lieu theo section (sang, trua, chieu, all)"""
    # Tim cac chi so cua LOGO DAO
    logo_dao_indices = []
    for i, title in enumerate(df_normalized['Ten_CT']):
        if pd.notna(title) and "logo dao" in unidecode(str(title).lower()):
            logo_dao_indices.append(i)
    
    # Nếu chọn "all", trả về từ LOGO DAO đầu tiên đến hết
    if section == "all":
        if len(logo_dao_indices) < 1:
            raise ValueError(f"Khong tim thay 'LOGO DAO' trong file HDPS.")
        return df_normalized.iloc[logo_dao_indices[0]:].reset_index(drop=True)
    
    if len(logo_dao_indices) < 3:
        raise ValueError(f"Khong tim thay du 3 'LOGO DAO' trong file HDPS. Chi tim thay {len(logo_dao_indices)} LOGO DAO.")
    
    # Xac dinh start va end index
    if section == "sang":
        start_index = logo_dao_indices[0]
        end_index = logo_dao_indices[1]
    elif section == "trua":
        start_index = logo_dao_indices[1]
        end_index = logo_dao_indices[2]
    elif section == "chieu":
        start_index = logo_dao_indices[2]
        end_index = len(df_normalized)
    else:
        raise ValueError(f"Section khong hop le: {section}")
    
    return df_normalized.iloc[start_index:end_index].reset_index(drop=True)


def convert_xls_to_xlsx(file_path):
    """Chuyen doi file .xls sang .xlsx bang MS Excel COM"""
    if not file_path.endswith('.xls'):
        return file_path
    
    # Tao ten file .xlsx
    xlsx_path = file_path + 'x'  # .xls -> .xlsx
    
    # Xoa file .xlsx cu neu ton tai (de ghi de)
    if os.path.exists(xlsx_path):
        try:
            os.remove(xlsx_path)
        except Exception as e:
            print(f"Khong the xoa file cu: {str(e)}")
    
    try:
        # Su dung Win32COM de mo Excel va save as
        try:
            import win32com.client as win32
            
            # Khoi tao Excel
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            excel.Visible = False
            excel.DisplayAlerts = False
            
            # Mo file .xls
            wb = excel.Workbooks.Open(os.path.abspath(file_path))
            
            # Save as .xlsx (FileFormat=51)
            wb.SaveAs(os.path.abspath(xlsx_path), FileFormat=51)
            
            # Dong file va Excel
            wb.Close(SaveChanges=False)
            excel.Quit()
            
            return xlsx_path
            
        except ImportError:
            raise ValueError("Khong tim thay thu vien pywin32. Vui long cai dat: pip install pywin32")
        except Exception as e:
            raise ValueError(f"Loi khi su dung Excel COM: {str(e)}")
        
    except Exception as e:
        print(f"Loi khi chuyen doi file: {str(e)}")
        raise ValueError(f"Khong the chuyen doi file .xls sang .xlsx. Vui long chuyen doi thu cong hoac cai dat MS Excel va pywin32")


def read_excel_safe(file_path, header=None):
    """Doc file Excel an toan, tu dong chuyen doi neu can"""
    original_path = file_path
    
    # Neu la file .xls, thu chuyen doi sang .xlsx
    if file_path.endswith('.xls'):
        try:
            file_path = convert_xls_to_xlsx(file_path)
        except Exception as e:
            raise ValueError(f"Khong the xu ly file .xls: {str(e)}\n\nVui long:\n1. Cai dat pywin32: pip install pywin32\n2. Hoac chuyen doi thu cong file .xls sang .xlsx")
    
    # Doc file voi openpyxl
    try:
        if header is not None:
            df = pd.read_excel(file_path, header=header, engine='openpyxl')
        else:
            df = pd.read_excel(file_path, engine='openpyxl')
        
        return df
    except Exception as e:
        raise ValueError(f"Khong the doc file {os.path.basename(original_path)}: {str(e)}")


def load_ad_config():
    """Load ad configuration from ad.json"""
    try:
        # Gia su file ad.json nam cung thu muc voi script
        script_dir = os.path.dirname(os.path.abspath(__file__))
        ad_json_path = os.path.join(script_dir, 'ad.json')
        
        with open(ad_json_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        print("Canh bao: Khong tim thay file ad.json")
        return {}
    except Exception as e:
        print(f"Loi khi doc file ad.json: {str(e)}")
        return {}

def load_replacements():
    """Load replacements configuration from replacements.json"""
    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        replacements_json_path = os.path.join(script_dir, 'replacements.json')
        
        with open(replacements_json_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        print("Canh bao: Khong tim thay file replacements.json")
        return {}
    except Exception as e:
        print(f"Loi khi doc file replacements.json: {str(e)}")
        return {}

def check_advertisements(channel, date_str, df_playlist_for_ad, ad_config):
    """Kiem tra cac file quang cao trong cac thu muc da chi dinh"""
    errors = []
    
    # Chon danh sach path theo kenh
    if channel == "3":
        ad_paths = CHANNEL_3_AD_PATHS
    elif channel == "4":
        ad_paths = CHANNEL_4_AD_PATHS
    else:
        return []
    
    # Chuyen doi date_str tu ddmmyyyy sang datetime
    try:
        date_obj = datetime.strptime(date_str, "%d%m%Y")
    except ValueError:
        errors.append("Loi: Khong the parse ngay tu ten file HDPS")
        return errors
    
    # Lay thoi gian bat dau va ket thuc cua playlist
    playlist_start_time = parse_time(df_playlist_for_ad.iloc[0]['Gio'])
    playlist_end_time = parse_time(df_playlist_for_ad.iloc[-1]['Gio'])
    
    # Nếu end_time < start_time nghĩa là đã qua ngày mới,chỉ tính tới 23:59:59
    if playlist_end_time < playlist_start_time:
        playlist_end_time = parse_time('23:59:59')
    
    # Tao dict de lookup nhanh cac file trong playlist
    # Key: normalized filename, Value: (original_filename, time_in_seconds)
    playlist_files = {}
    for idx, row in df_playlist_for_ad.iterrows():
        file_name = str(row['Ten_CT']).strip()
        file_name_normalized = normalize_title(file_name)
        file_time = parse_time(row['Gio'])
        playlist_files[file_name_normalized] = (file_name, file_time)
        #print (f"{file_time} : {file_name_normalized}")
    
    # Duyet qua tung path
    for path_template in ad_paths:
        # Format path voi ngay thang
        formatted_path = date_obj.strftime(path_template)
        
        # Kiem tra xem thu muc co ton tai khong
        if not os.path.exists(formatted_path):
            #errors.append(f"\nKhong tim thay: {formatted_path}")
            continue
        
        # Lay danh sach tat ca cac file trong thu muc
        try:
            files = os.listdir(formatted_path)
        except Exception as e:
            errors.append(f"\nLoi khi doc thu muc {formatted_path}: {str(e)}")
            continue
        
        # Duyet qua tung file
        for file_name in files:
            # Loại bỏ extension để so sánh
            file_name_without_ext = os.path.splitext(file_name)[0]
        
            # Kiem tra xem file co chua bat ky key nao trong ad_config khong
            for ad_key, ad_time_range in ad_config.items():
                # Chi xu ly cac key cua kenh tuong ung
                if not ad_key.startswith(channel):
                    continue
                
                # Kiem tra xem ad_key co trong ten file khong (khong phan biet hoa thuong)
                if ad_key.lower() not in file_name.lower():
                    continue
                
                # Lay start_time va end_time tu ad_config
                try:
                    ad_start_time = parse_time(ad_time_range['start_time'])
                    ad_end_time = parse_time(ad_time_range['end_time'])
                except Exception as e:
                    errors.append(f"\nLoi khi parse thoi gian cho {ad_key}: {str(e)}")
                    continue
                
                # Kiem tra xem thoi gian co nam trong khoang cua playlist khong
                if ad_end_time < playlist_start_time or ad_start_time > playlist_end_time:
                    # Thoi gian quang cao nam ngoai khoang playlist, bo qua
                    #print(f"playlist start time {format_time(playlist_start_time)}")
                    #print(f"playlist end time {format_time(playlist_end_time)}")
                    #print(f"ad start time {format_time(ad_start_time)}")
                    #print(f"ad end time {format_time(ad_end_time)}")
                    continue
                
                # Kiem tra xem file co trong playlist khong
                file_name_normalized = normalize_title(file_name_without_ext)
                
                found = False
                match = False
                matched_file_time = None
                for playlist_key, (original_name, file_time) in playlist_files.items():
                    # Kiem tra xem ten file co match voi entry trong playlist khong
                    #if file_name_normalized in playlist_key or playlist_key in file_name_normalized:
                    if file_name_normalized == playlist_key:
                        found = True
                        matched_file_time = file_time
                        # Kiem tra xem thoi gian co nam trong khoang start_time, end_time khong
                        if ad_start_time <= file_time <= ad_end_time:
                            match = True
                            break
                
                if not found:
                    errors.append(f"\n[QUANG CAO LOI] File: {file_name}")
                    errors.append(f"   Khong tim thay trong playlist")
                elif not match:
                    errors.append(f"\n[QUANG CAO LOI] File: {file_name}")
                    errors.append(f"   Thoi khoang khong khop - thoi khoang chuan: {format_time(ad_start_time)} - {format_time(ad_end_time)}")
                    if matched_file_time:
                        errors.append(f"   Thoi khoang thuc te: {format_time(matched_file_time)}")
    
    return errors


def check_hdps_and_playlist(excel_path_hdps, excel_path_playlist, date, section, channel):
    """Ham chinh de so sanh hai file Excel"""
    
    # Load ad configuration
    ad_config = load_ad_config()
    
    # Load replacements configuration
    global REPLACEMENTS
    REPLACEMENTS = load_replacements()
    
    # Doc file Excel HDPS
    try:
        df_hdps_raw = read_excel_safe(excel_path_hdps, header=1)
        
        # Chuan hoa DataFrame HDPS
        df_hdps = normalize_hdps_dataframe(df_hdps_raw)
    except Exception as e:
        raise ValueError(f"Loi khi doc file HDPS: {str(e)}")
    
    # Doc file Excel Playlist
    try:
        df_playlist_raw = read_excel_safe(excel_path_playlist, header=0)
        
        # Chuan hoa DataFrame Playlist
        df_playlist, df_playlist_for_ad, part_order_errors  = normalize_playlist_dataframe(df_playlist_raw)
    except Exception as e:
        raise ValueError(f"Loi khi doc file Playlist: {str(e)}")
    
    # Loc du lieu theo section
    df_section = get_section_data(df_hdps, section)
    print(df_section)
    print(df_playlist)
    #print(df_playlist_for_ad)
    
    result_lines = []
    result_lines.append(f"=== KET QUA SO SANH - SECTION: {section.upper()} ===\n")
    result_lines.append(f"Tong so dong HDPS (section {section}): {len(df_section)}")
    result_lines.append(f"Tong so dong Playlist: {len(df_playlist)}\n")
    
    errors = []
    match_count = 0
    
    # So sanh tung dong
    min_len = min(len(df_section), len(df_playlist))
    
    for idx in range(min_len):
        hdps_row = df_section.iloc[idx]
        playlist_row = df_playlist.iloc[idx]
        
        # Kiem tra xem co phai la new_row khong
        is_new_row = hdps_row.get('is_new_row', False)
        
        # Kiem tra xem co phai la dau_phim khong
        is_dau_phim = hdps_row.get('is_dau_phim', False)
        
        # So sanh thoi gian
        hdps_time = parse_time(hdps_row['Gio'])
        playlist_time = parse_time(playlist_row['Gio'])
        time_diff = abs(hdps_time - playlist_time)
        
        # So sanh thoi luong
        hdps_duration = parse_time(hdps_row['Thoi_luong'])
        playlist_duration = parse_time(playlist_row['Thoi_luong'])
        duration_diff = abs(hdps_duration - playlist_duration)
        
        # So sanh ten chuong trinh
        hdps_title = normalize_title(hdps_row['Ten_CT'])
        playlist_title = normalize_title(playlist_row['Ten_CT'])
        playlist_title_base = playlist_row['Ten_CT']
        
        # Tinh do tuong dong
        if hdps_title and playlist_title:
            title_ratio = fuzz.ratio(hdps_title, playlist_title)
        else:
            title_ratio = 0
        
        # Kiem tra dieu kien
        is_ok = True
        error_details = []
        
        # CHI KIEM TRA TIME VA DURATION NEU KHONG PHAI LA NEW_ROW
        if not is_new_row and not is_dau_phim:
            if time_diff > 180:
                is_ok = False
                error_details.append(f"  Gio: HDPS={format_time(hdps_time)}, Playlist={format_time(playlist_time)}, chenh={time_diff:.1f}s")
            if duration_diff > 10:
                is_ok = False
                error_details.append(f"  Thoi luong: HDPS={format_time(hdps_duration)}, Playlist={format_time(playlist_duration)}, chenh={duration_diff:.1f}s")
        elif is_new_row and is_dau_phim:
            if time_diff > 180:
                is_ok = False
                error_details.append(f"  Gio: HDPS={format_time(hdps_time)}, Playlist={format_time(playlist_time)}, chenh={time_diff:.1f}s")
            
        # LUON KIEM TRA TEN CHUONG TRINH
        # Kiem tra neu ca hai deu chua "het tap" thi dung nguong 100%
        if "het tap" in hdps_title and "het tap" in playlist_title:
            if title_ratio < 100:  # Yeu cau chinh xac 100%
                is_ok = False
                error_details.append(f"  Ten CT: match={title_ratio}% (< 100%)")
                error_details.append(f"  HDPS: {hdps_title[:100]}")
                error_details.append(f"  Playlist: {playlist_title[:100]}")
        else:
            if title_ratio < MIN_MATCH_SCORE:
                is_ok = False
                error_details.append(f"  Ten CT: match={title_ratio}% (< {MIN_MATCH_SCORE}%)")
                error_details.append(f"  HDPS: {hdps_title[:100]}")
                error_details.append(f"  Playlist: {playlist_title[:100]}")
        
        if is_ok:
            match_count += 1
        else:
            errors.append(f"\n[PLAYLIST LOI] File: {playlist_title_base}")
            for detail in error_details:
                errors.append(f"   {detail}")
    
    # Kiem tra neu so dong khac nhau
    if len(df_section) > len(df_playlist):
        for idx in range(len(df_playlist), len(df_section)):
            errors.append(f"\nHang {idx + 1}: HDPS co them dong nhung Playlist da het")
    elif len(df_playlist) > len(df_section):
        for idx in range(len(df_section), len(df_playlist)):
            errors.append(f"\nHang {idx + 1}: Playlist co them dong nhung HDPS da het")
    
    # Kiem tra quang cao neu co ad_config
    if ad_config:
        ad_errors = check_advertisements(channel, date, df_playlist_for_ad, ad_config)
        if ad_errors:
            errors.append(f"\n\n{'='*60}")
            errors.append("KIEM TRA QUANG CAO:")
            errors.append(f"{'='*60}")
            errors.extend(ad_errors)
            
    # Them loi thu tu part vao errors (sau phan kiem tra quang cao)
    if part_order_errors:
        errors.append(f"\n\n{'='*60}")
        errors.append("LOI THU TU PART:")
        errors.append(f"{'='*60}")
        for err in part_order_errors:
            errors.append(f"\n[LOI THU TU PART] {err['title']}")
            #errors.append(f"   Dong {err['row']}: Mong doi Part {err['expected']}, nhung nhan duoc Part {err['actual']}")        
    
    # Tong hop ket qua
    #result_lines.append(f"\n{'='*60}")
    #result_lines.append(f"So dong khop: {match_count}/{min_len}")
    #result_lines.append(f"So dong loi: {len([e for e in errors if 'PLAYLIST LOI' in e or 'QUANG CAO LOI' in e])}")
    #result_lines.append(f"{'='*60}\n")
    
    if errors:
        result_lines.append(f"\n\n{'='*60}")
        result_lines.append("KIEM TRA PLAYLIST:")
        result_lines.append(f"{'='*60}")
        result_lines.extend(errors)
    else:
        result_lines.append("TAT CA CAC DONG DEU KHOP!")
    
    return "\n".join(result_lines)


def main():
    root = tk.Tk()
    app = CheckPlaylistApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()