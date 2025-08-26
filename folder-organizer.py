import os, shutil

folder = "Downloads"
for file in os.listdir(folder):
    if file.endswith((".jpg", ".png")):
        shutil.move(os.path.join(folder, file), os.path.join(folder, "Images"))
    elif file.endswith(".pdf"):
        shutil.move(os.path.join(folder, file), os.path.join(folder, "PDFs"))