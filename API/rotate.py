import numpy as np
from PIL import Image
import re
import os
import pandas as pd
import ast
from tqdm import tqdm

def rotate_to_vertical(img_path, points):
    image = Image.open(img_path).convert("RGB")  
    width_bbox = np.linalg.norm(points[0] - points[1])
    height_bbox = np.linalg.norm(points[0] - points[3])
    width_real, height_real = image.size

    if np.abs(np.round(width_bbox) - np.round(width_real)) >= 1 and np.abs(np.round(height_bbox) - np.round(height_real)) >= 1 :
        rotated_image = image.rotate(-90, expand=True, fillcolor=(255, 255, 255))
        return rotated_image, 90
    else:
        return image, 0


def handle_rotate(path_folder, path_dict = "data/upload/samples.xlsx"):
    df = pd.read_excel(path_dict)
    images = os.listdir(path_folder)
    for image in tqdm(images, desc="watting for rotate: ", unit="image"):
        if not image.lower().endswith(('png', 'jpg', 'jpeg', 'bmp', 'gif')):
            continue
        image_path = os.path.join(path_folder, image)
        bbox_str = df[df["Img_Box_ID"] == image]["Img_Box_Coordinate"].values
        if len(bbox_str) == 0:
            print(f"ảnh: {image} không tìm thấy trong {path_folder}")
            continue
        match = re.search(r"\[\[.*?\]\]",bbox_str[0])
        if match:
            try:
                bbox_array = np.array(ast.literal_eval(match.group()), np.float32)
                img_rotate, _ = rotate_to_vertical(image_path, bbox_array)
                os.remove(image_path)
                img_rotate.save(image_path)
            except Exception as e:
                print(f"lỗi xủ lý ảnh {image}: {e}")
        else:
            print(f"Không tìm thấy bbox hợp lệ cho: {image}")


