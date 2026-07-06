# Tool gop BOM

Cong cu dinh dang file BOM Excel de in di cho.

## Cach chay nhanh

Mo PowerShell hoac Command Prompt tai thu muc nay:

```powershell
cd "C:\Users\Cua\Desktop\proj\Tool gộp"
```

Chay web local:

```powershell
.\start_web.bat
```

Sau do mo trinh duyet:

```text
http://127.0.0.1:8100
```

Neu muon cho may khac trong cung mang LAN truy cap:

```powershell
.\start_lan.bat
```

## Chay bang Python

May can co Python va `openpyxl`.

```powershell
pip install -r requirements.txt
python main.py --host 127.0.0.1 --port 8100
```

Khong tu mo trinh duyet:

```powershell
python main.py --host 127.0.0.1 --port 8100 --no-browser
```

## Xu ly hang loat

Bo file `.xlsx`, `.xlsm`, hoac `.csv` vao thu muc `Excel`, roi chay:

```powershell
python main.py --batch
```

Ket qua se nam trong thu muc `Đã gộp`. Moi file dau vao se xuat cac dinh dang:

- `Format 1`: Bang A4 doc 13 cot + dong ngay, `- da dinh dang.xlsx`
- `Format 2`: Theo mau `Mẫu/Mẫu format 2.xlsx`, `- format 2.xlsx`
- `Duyet dinh muc`: Bang A4 ngang de duyet thu vien BOM, gop cac mon an lien ke, `- duyet dinh muc.xlsx`

## Luu y

- Nen chay `start_web.bat` o thu muc goc nay de dung code moi nhat.
- Cac file `.exe` trong `dist` chi dung khi da build lai sau khi sua code.
- Neu port `8100` dang ban, co the chay port khac:

```powershell
python main.py --port 8110
```
