#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
创建寻拟邮件工具的图标
使用PIL库生成一个现代化的邮件图标
"""

from PIL import Image, ImageDraw, ImageFont
import os


def create_icon():
    """创建图标"""

    # 创建多个尺寸的图标
    sizes = [256, 128, 64, 48, 32, 16]
    images = []

    for size in sizes:
        # 创建画布
        img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)

        # 计算缩放比例
        scale = size / 256

        # 背景渐变圆角矩形
        padding = int(10 * scale)
        corner_radius = int(30 * scale)

        # 绘制圆角矩形背景(蓝色渐变)
        draw.rounded_rectangle(
            [padding, padding, size - padding, size - padding],
            radius=corner_radius,
            fill=(52, 152, 219, 255),  # 蓝色 #3498db
            outline=None
        )

        # 绘制信封
        envelope_margin = int(40 * scale)
        envelope_top = int(80 * scale)
        envelope_bottom = int(190 * scale)
        envelope_left = int(40 * scale)
        envelope_right = int(216 * scale)

        # 信封主体(白色)
        envelope_points = [
            (envelope_left, envelope_top + int(20 * scale)),
            (envelope_left, envelope_bottom),
            (envelope_right, envelope_bottom),
            (envelope_right, envelope_top + int(20 * scale))
        ]
        draw.polygon(envelope_points, fill=(255, 255, 255, 255))

        # 信封盖(三角形)
        flap_points = [
            (envelope_left, envelope_top + int(20 * scale)),
            (int(128 * scale), envelope_top - int(10 * scale)),
            (envelope_right, envelope_top + int(20 * scale))
        ]
        draw.polygon(flap_points, fill=(236, 240, 241, 255))  # 浅灰色

        # 信封边框线条
        draw.line([envelope_left, envelope_top + int(20 * scale),
                   int(128 * scale), envelope_top - int(10 * scale)],
                  fill=(189, 195, 199, 255), width=int(2 * scale))
        draw.line([int(128 * scale), envelope_top - int(10 * scale),
                   envelope_right, envelope_top + int(20 * scale)],
                  fill=(189, 195, 199, 255), width=int(2 * scale))

        # 添加装饰线条(信封上的横线)
        if size >= 64:
            line_y_start = int(110 * scale)
            line_spacing = int(15 * scale)
            line_width = int(2 * scale)

            for i in range(3):
                y = line_y_start + i * line_spacing
                draw.line(
                    [envelope_left + int(20 * scale), y,
                     envelope_right - int(20 * scale), y],
                    fill=(52, 152, 219, 200),  # 半透明蓝色
                    width=line_width
                )

        # 添加"AI"标识(右上角)
        if size >= 48:
            ai_size = int(40 * scale)
            ai_x = size - padding - ai_size
            ai_y = padding

            # AI圆形背景
            draw.ellipse(
                [ai_x, ai_y, ai_x + ai_size, ai_y + ai_size],
                fill=(231, 76, 60, 255),  # 红色 #e74c3c
                outline=None
            )

            # AI文字
            if size >= 64:
                try:
                    # 尝试使用系统字体
                    font_size = int(20 * scale)
                    font = ImageFont.truetype("arial.ttf", font_size)
                except:
                    font = ImageFont.load_default()

                text = "AI"
                # 获取文本边界框
                bbox = draw.textbbox((0, 0), text, font=font)
                text_width = bbox[2] - bbox[0]
                text_height = bbox[3] - bbox[1]

                text_x = ai_x + (ai_size - text_width) // 2
                text_y = ai_y + (ai_size - text_height) // 2 - bbox[1]

                draw.text((text_x, text_y), text, fill=(255, 255, 255, 255), font=font)

        images.append(img)

    return images


def save_icons():
    """保存图标文件"""
    print("正在生成寻拟邮件工具图标...")

    images = create_icon()

    # 保存为ICO文件(Windows图标)
    ico_path = "icon.ico"
    images[0].save(ico_path, format='ICO', sizes=[(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)])
    print(f"[OK] 已生成: {ico_path}")

    # 保存为PNG文件(各种尺寸)
    sizes = [256, 128, 64, 48, 32, 16]
    for i, size in enumerate(sizes):
        png_path = f"icon_{size}x{size}.png"
        images[i].save(png_path, format='PNG')
        print(f"[OK] 已生成: {png_path}")

    # 保存一个大尺寸PNG用于展示
    images[0].save("icon.png", format='PNG')
    print(f"[OK] 已生成: icon.png (主图标)")

    print("\n图标生成完成!")
    print("\n使用方法:")
    print("1. icon.ico - 用于Windows应用程序图标")
    print("2. icon.png - 用于文档、网站等展示")
    print("3. icon_*x*.png - 不同尺寸的PNG图标")


if __name__ == "__main__":
    save_icons()
