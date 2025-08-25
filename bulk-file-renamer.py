import os

folder = "path/to/folder"
for i, file in enumerate(os.listdir(folder)):
    ext = file.split('.')[-1]
    os.rename(os.path.join(folder, file), os.path.join(folderm f"file_{i}.{ext}"))