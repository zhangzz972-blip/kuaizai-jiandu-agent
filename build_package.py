#!/usr/bin/env python3
"""Build distributable package for 快哉荐读 agent v1.0"""
import os, shutil, zipfile

DIST = "dist/快哉荐读agent_v1.0"

# Clean and recreate dist
if os.path.exists(DIST):
    shutil.rmtree(DIST)
os.makedirs(DIST)
os.makedirs(f"{DIST}/templates", exist_ok=True)

# Source files
FILES = [
    "app.py",
    "jiandu_deepseek.py",
    "templates/index.html",
    "启动.bat",
    "setup.bat",
    "README.md",
    ".env.example",
    ".gitignore",
]

for f in FILES:
    src = f
    dst = f"{DIST}/{f}"
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    if os.path.exists(src):
        shutil.copy2(src, dst)

# Copy data files if present
for df in ["参考咨询书库馆藏清单.xlsx", "馆标黑版.png"]:
    if os.path.exists(df):
        shutil.copy2(df, f"{DIST}/{df}")
        print(f"  [include] {df}")
    else:
        print(f"  [skip] {df} not found — user must provide")

# Create zip
zipname = "快哉荐读agent_v1.0_便携版.zip"
zip_path = f"dist/{zipname}"
with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
    for root, dirs, files in os.walk(DIST):
        for fn in files:
            full = os.path.join(root, fn)
            arc = os.path.relpath(full, DIST)
            zf.write(full, arc)

size_mb = os.path.getsize(zip_path) / 1024 / 1024
print(f"\nDone: {zip_path} ({size_mb:.1f} MB)")

# Print usage
print(f"""
分发说明：
  1. 将 {zipname} 发给使用者
  2. 解压到任意目录
  3. 确保安装 Python 3.10+
  4. 双击 setup.bat（仅首次）
  5. 双击 启动.bat（每次使用）
""")
