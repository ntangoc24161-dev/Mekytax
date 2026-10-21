---
type: source
created: 2026-10-15
updated: 2026-10-15
---
# Gắn Google Drive vào Javis trên VPS bằng rclone

## Cách nhanh (không SSH): authorize trên Windows, dán token vào chat

Javis đã có rclone `v1.75.2` tại `/data/state/bin/rclone`, file cấu hình `/data/state/rclone/rclone.conf`. Remote tên `gdrive`, scope My Drive. **Chưa có token.**

Google không nhận mail/mật khẩu trong chat. Anh đăng nhập trên trình duyệt máy Windows:

1. Cài rclone trên Windows (PowerShell Admin): `winget install Rclone.Rclone`
2. Chạy đúng lệnh:

```powershell
rclone authorize "drive" "eyJzY29wZSI6ImRyaXZlIn0"
```

3. Chọn đúng Gmail đang chứa `HS CONG AN-2025-2026`, cho phép.
4. Cửa sổ in một khối JSON. Dán nguyên khối đó vào chat Javis (bắt đầu bằng `{`).

Em gắn token rồi kéo `Bao_cao_cong_van.xlsx`. Không gửi mật khẩu Gmail.

Nếu Google báo client_id rclone dùng chung đã ngừng (2026): tạo client_id riêng theo https://rclone.org/drive/#making-your-own-client-id rồi nhắn lại.

---

Làm **trên máy chủ VPS qua SSH** nếu muốn gắn ổ fuse `/mnt/i` như `I:\`. Container không có `/dev/fuse`, không tự mount được.

Mục tiêu fuse: Drive `My Drive` hiện trong Javis tại `/mnt/i/`, tương ứng `I:\My Drive\` trên Windows.

Ổ `D:\HSO CONG AN FILE MD` là đĩa máy Windows, rclone **không** lấy được.

## 0. Cảnh báo trước khi sửa compose

Nếu stack Javis đang kéo YAML từ URL GitHub (`docker-compose.hostinger.yml`), lần **Redeploy / Cập nhật** có thể **xóa** dòng volume vừa thêm. Sau khi gắn Drive, giữ bản compose đã sửa trên VPS (không deploy lại từ URL trần), hoặc nhớ dán lại dòng `/mnt/gdrive:/mnt/i:rshared`.

Không sửa cùng một file Excel vừa trên Google Drive Desktop (máy Windows) vừa trên VPS cùng lúc.

## 1. SSH vào VPS

Trên Hostinger: hPanel → VPS → SSH. User cần `sudo`.

```bash
whoami
docker ps --format 'table {{.Names}}\t{{.Image}}\t{{.Status}}'
ls /mnt/gdrive 2>/dev/null || echo "chua co /mnt/gdrive"
```

Tên container mặc định: `javis`.

## 2. Cài rclone và FUSE trên HOST

```bash
sudo apt update
sudo apt install -y fuse3 unzip curl
curl https://rclone.org/install.sh | sudo bash
rclone version
```

Cho phép mount `--allow-other`:

```bash
echo user_allow_other | sudo tee -a /etc/fuse.conf
```

## 3. Đăng nhập Google Drive (máy VPS không có trình duyệt)

```bash
rclone config
```

Chọn lần lượt:

- `n` (new remote)
- tên: `gdrive`
- storage: `drive` (Google Drive)
- `client_id` / `client_secret`: Enter (để trống)
- scope: `1` (Full access) - cần ghi Excel
- `root_folder_id`: Enter
- service account: Enter
- auto config: `n` (headless)

rclone in một **link**. Mở link đó trên máy Windows đã đăng nhập đúng Google Drive có thư mục `HS CONG AN-2025-2026`, cho phép, rồi **dán token** lại cửa sổ SSH.

Kiểm tra thấy đúng thư mục:

```bash
rclone lsf gdrive: --max-depth 1
rclone lsf "gdrive:HS CONG AN-2025-2026/HS CUNG CAP CA 2026" | head
```

Phải thấy `Bao_cao_cong_van.xlsx`. Không thấy thì:

```bash
rclone ls gdrive: --include "*Bao_cao_cong_van*" --max-depth 5
```

## 4. Mount bằng systemd (sống sau khi reboot)

```bash
sudo mkdir -p /mnt/gdrive
sudo chown 10001:10001 /mnt/gdrive
```

`10001` là user trong container Javis.

Tạo service:

```bash
sudo tee /etc/systemd/system/rclone-gdrive.service > /dev/null << 'EOF'
[Unit]
Description=rclone mount Google Drive (My Drive)
After=network-online.target
Wants=network-online.target

[Service]
Type=notify
User=root
ExecStart=/usr/bin/rclone mount gdrive: /mnt/gdrive \
  --allow-other \
  --uid 10001 \
  --gid 10001 \
  --umask 002 \
  --dir-cache-time 15s \
  --vfs-cache-mode writes \
  --vfs-cache-max-age 1h \
  --vfs-write-back 10s
ExecStop=/bin/fusermount3 -u /mnt/gdrive
Restart=on-failure
RestartSec=10

[Install]
WantedBy=multi-user.target
EOF

sudo systemctl daemon-reload
sudo systemctl enable --now rclone-gdrive.service
sudo systemctl status rclone-gdrive.service --no-pager
ls "/mnt/gdrive/HS CONG AN-2025-2026/HS CUNG CAP CA 2026/Bao_cao_cong_van.xlsx"
```

`--vfs-cache-mode writes` để ghi file `.xlsx` không bị Drive từ chối.

## 5. Gắn vào Docker compose rồi restart Javis

Trong Docker Manager (hoặc file compose trên VPS), **thêm một dòng** vào `volumes` của service `javis`, giữ nguyên các volume cũ:

```yaml
    volumes:
      - javis-data:/data
      - javis-brains:/brains
      - claude-auth:/home/javis/.claude
      - /mnt/gdrive:/mnt/i:rshared
```

Rồi **Restart** container `javis` (không cần Rebuild image).

Nếu YAML không nhận `:rshared`, thử:

```yaml
      - /mnt/gdrive:/mnt/i
```

Thứ tự bắt buộc: **rclone mount đã chạy** rồi mới start/restart Javis. Nếu Javis lên trước, `/mnt/i` sẽ là thư mục trống.

```bash
sudo systemctl start rclone-gdrive.service
docker restart javis
```

## 6. Kiểm tra từ trong Javis

Nhắn lại trên chat: `kiểm tra /mnt/i`. Đường dẫn sổ chính lúc đó là:

`/mnt/i/HS CONG AN-2025-2026/HS CUNG CAP CA 2026/Bao_cao_cong_van.xlsx`

Tương ứng Windows: `I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\Bao_cao_cong_van.xlsx`

`DB NNT.xls` nếu nằm gốc My Drive: `/mnt/i/DB NNT.xls`.

## 7. Nếu `/mnt/i` trống sau khi restart

FUSE vào Docker đôi khi bind nhầm thư mục rỗng. Phương án chắc hơn: **sync một thư mục thường** (không phải fuse) rồi bind thư mục đó.

```bash
sudo mkdir -p "/opt/javis-drive/HS CONG AN-2025-2026"
sudo chown -R 10001:10001 /opt/javis-drive
sudo rclone sync "gdrive:HS CONG AN-2025-2026" "/opt/javis-drive/HS CONG AN-2025-2026" --progress
```

Compose:

```yaml
      - /opt/javis-drive:/mnt/i
```

Sau khi Javis ghi Excel, đẩy ngược lên Drive:

```bash
sudo rclone copy "/opt/javis-drive/HS CONG AN-2025-2026" "gdrive:HS CONG AN-2025-2026" --update --progress
```

Cách này không live như ổ `I:`, nhưng ít gãy trên Hostinger.

## Gỡ mount

```bash
sudo systemctl disable --now rclone-gdrive.service
sudo fusermount3 -u /mnt/gdrive
```

Xóa dòng `/mnt/gdrive:/mnt/i` trong compose rồi restart Javis.
