# Deploy tool dinh dang BOM

## Cach dung tren may hien tai

Chay:

```bat
start_web.bat
```

Trang web se tu mo tai `http://127.0.0.1:8000`.

## Dong goi cho may khac

Tren may build co Python va Internet, chay:

```bat
build_exe.bat
```

Sau khi build xong, copy toan bo thu muc `dist\BomFormatterWeb` sang may khac. Tren may nhan, chay:

```bat
start_web.bat
```

May nhan khong can cai Python.

## Cho nhieu may trong cung mang LAN dung chung

Tren may chay server, mo:

```bat
start_lan.bat
```

Xem dong `IPv4 Address`, sau do may khac trong cung mang truy cap:

```text
http://<IPv4-cua-may-server>:8000
```

Neu Windows Firewall hoi quyen, chon allow cho Python hoac `BomFormatterWeb.exe`.
