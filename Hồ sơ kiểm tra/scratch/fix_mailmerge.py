import zipfile, os, shutil, tempfile

docx_path = r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\1. TD\12. TT_BIEN_HOA_3603883505\0. BIEN BAN\TONG HOP.docx"
backup_path = r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\1. TD\12. TT_BIEN_HOA_3603883505\0. BIEN BAN\TONG HOP_backup.docx"

if not os.path.exists(backup_path):
    shutil.copyfile(docx_path, backup_path)

temp_dir = tempfile.mkdtemp()
with zipfile.ZipFile(docx_path, 'r') as z:
    z.extractall(temp_dir)

settings_path = os.path.join(temp_dir, 'word', 'settings.xml')
with open(settings_path, 'r', encoding='utf-8') as f:
    s = f.read()

# 1. Update Extended Properties to include Excel 12.0 Macro for .xlsm files
old_prop = 'Extended Properties=&quot;HDR=YES;IMEX=1;&quot;;'
new_prop = 'Extended Properties=&quot;Excel 12.0 Macro;HDR=YES;IMEX=1;&quot;;'
s_new = s.replace(old_prop, new_prop)

# 2. Match table val with query
# In settings.xml: <ns0:query ns0:val="SELECT * FROM `01_DA_BAN_HANH$`" />
# In odso: <ns0:table ns0:val="01_DA_BAN_HANH$" /> -> must be `01_DA_BAN_HANH$`
s_new = s_new.replace('table ns0:val="01_DA_BAN_HANH$"', 'table ns0:val="`01_DA_BAN_HANH$`"')

with open(settings_path, 'w', encoding='utf-8') as f:
    f.write(s_new)

out_zip = tempfile.mktemp(suffix='.docx')
with zipfile.ZipFile(out_zip, 'w', zipfile.ZIP_DEFLATED) as z_out:
    for root, dirs, files in os.walk(temp_dir):
        for file in files:
            full_p = os.path.join(root, file)
            rel_p = os.path.relpath(full_p, temp_dir)
            z_out.write(full_p, rel_p)

shutil.copyfile(out_zip, docx_path)
os.remove(out_zip)
shutil.rmtree(temp_dir)
print("SUCCESS")
