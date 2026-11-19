# Chứng minh OAuth/Drive — 19/11/2026 (lần sửa kiểm chứng)

## Phản hồi bị chê
`Failed to authenticate: OAuth session expired and could not be refreshed`

## Cách xác minh (không dùng OAuth engine Claude)
1. Hub Javis `POST /hub/mcp` với `Authorization: Bearer <hub_token>` (file vault `.grok/config.toml`).
2. Header `X-Javis-Vault` = URL-encode của `/brains/Cung cấp thông tin công an`.
3. `javis_connections` → Composio trạng thái **ổn**.
4. `composio__COMPOSIO_MANAGE_CONNECTIONS` action=list toolkit=googledrive → `googledrive_ralph-urbane` **ACTIVE**, mail `cungcapthongtin.tcs3tdn@gmail.com`.
5. `COMPOSIO_MULTI_EXECUTE_TOOL` với `account` là **sibling** của `tool_slug` (không nhét trong `arguments`).
6. `GOOGLEDRIVE_FIND_FILE` + `GOOGLEDRIVE_DOWNLOAD_FILE` → success 4/4.

## Kết quả dữ liệu
- CV MOI trống.
- Folder 295/296 đủ 4 file mỗi folder (2 PDF + 2 md).
- Sổ sống dòng 896–898 đúng; DASHBOARD còn 2 chart; helper không `==`.
- File chứng: `Bao_cao_cong_van_drive_xac_minh_oauth_fix_2026-11-19.xlsx`.
