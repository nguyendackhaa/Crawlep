#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Module xử lý đặt tên file ảnh theo logic từ code JavaScript
"""

import re
import pandas as pd
import os

class ImageNamingProcessor:
    def __init__(self, domain="https://example.com/product/", image_base="https://cdn.example.com/images/"):
        """
        Khởi tạo processor với domain và image base URL
        
        Args:
            domain (str): Domain cho product URL
            image_base (str): Base URL cho image
        """
        self.domain = domain
        self.image_base = image_base
    
    def standardize(self, text):
        """
        Chuẩn hóa chuỗi theo logic từ JavaScript
        
        Args:
            text (str): Chuỗi cần chuẩn hóa
            
        Returns:
            str: Chuỗi đã chuẩn hóa
        """
        if not text:
            return ""
        
        # Thay ký tự không hợp lệ bằng '-'
        standardized = re.sub(r'[\\/:*?"<>|,=\s]', '-', text)
        
        # Loại bỏ gạch ngang trùng
        standardized = re.sub(r'-+', '-', standardized)
        
        # Chỉ giữ lại a-z, A-Z, 0-9, -, _
        standardized = re.sub(r'[^a-zA-Z0-9\-_]', '', standardized)
        
        # Loại bỏ gạch đầu/cuối
        standardized = re.sub(r'^-+|-+$', '', standardized)
        
        return standardized
    
    def process_product_code(self, code):
        """
        Xử lý mã sản phẩm để tạo slug và image name
        
        Args:
            code (str): Mã sản phẩm gốc
            
        Returns:
            tuple: (slug, image_name, had_addon_kit)
        """
        if not code:
            return "", "", False
        
        # Kiểm tra có add-on kit không
        had_addon_kit = bool(re.search(r'add[\s\-]*on[\s\-]*kit', code, re.IGNORECASE))
        
        # Làm sạch chuỗi, loại bỏ ghi chú coating, addon...
        clean_code = code
        clean_code = re.sub(r'\(with special coating\)', '', clean_code, flags=re.IGNORECASE)
        clean_code = re.sub(r'\[with special coating\]', '', clean_code, flags=re.IGNORECASE)
        clean_code = re.sub(r'add[\s\-]*on[\s\-]*kit', '', clean_code, flags=re.IGNORECASE)
        clean_code = clean_code.strip()
        
        # Tạo slug (lowercase) và image name (uppercase)
        slug = self.standardize(clean_code.lower())
        if had_addon_kit:
            slug += "-adk"
        
        image_name = self.standardize(clean_code.upper())
        
        return slug, image_name, had_addon_kit
    
    def generate_urls(self, slug, image_name):
        """
        Tạo product URL và image URL
        
        Args:
            slug (str): Slug sản phẩm
            image_name (str): Tên ảnh
            
        Returns:
            tuple: (product_url, image_url)
        """
        if not slug:
            return "", ""
        
        product_url = self.domain + slug
        image_url = self.image_base + image_name + ".webp"
        
        return product_url, image_url
    
    def process_excel_file(self, input_file, output_file=None, start_row=2, input_col=1):
        """
        Xử lý file Excel với logic đặt tên
        
        Args:
            input_file (str): Đường dẫn file Excel input
            output_file (str): Đường dẫn file Excel output (nếu None thì ghi đè)
            start_row (int): Dòng bắt đầu (mặc định 2, bỏ qua header)
            input_col (int): Cột chứa mã sản phẩm (mặc định 1 = cột B)
            
        Returns:
            str: Đường dẫn file output
        """
        try:
            # Đọc file Excel
            df = pd.read_excel(input_file)
            
            # Lấy cột mã sản phẩm (input_col - 1 vì pandas index từ 0)
            product_codes = df.iloc[:, input_col - 1].tolist()
            
            # Xử lý từng mã sản phẩm
            product_urls = []
            image_urls = []
            
            for code in product_codes:
                slug, image_name, had_addon = self.process_product_code(str(code))
                product_url, image_url = self.generate_urls(slug, image_name)
                
                product_urls.append(product_url)
                image_urls.append(image_url)
            
            # Thêm cột mới vào DataFrame
            df['Product_URL'] = product_urls
            df['Image_URL'] = image_urls
            
            # Lưu file
            if output_file is None:
                output_file = input_file.replace('.xlsx', '_processed.xlsx')
            
            df.to_excel(output_file, index=False)
            
            print(f"✅ Đã xử lý thành công: {output_file}")
            print(f"📊 Tổng số dòng: {len(df)}")
            print(f"🔗 Product URLs: {len([u for u in product_urls if u])}")
            print(f"🖼️ Image URLs: {len([u for u in image_urls if u])}")
            
            return output_file
            
        except Exception as e:
            print(f"❌ Lỗi khi xử lý file Excel: {str(e)}")
            return None
    
    def process_single_code(self, code):
        """
        Xử lý một mã sản phẩm đơn lẻ
        
        Args:
            code (str): Mã sản phẩm
            
        Returns:
            dict: Thông tin kết quả
        """
        slug, image_name, had_addon = self.process_product_code(code)
        product_url, image_url = self.generate_urls(slug, image_name)
        
        return {
            'original_code': code,
            'clean_code': slug,
            'image_name': image_name,
            'had_addon_kit': had_addon,
            'product_url': product_url,
            'image_url': image_url,
            'filename': image_name + ".webp" if image_name else ""
        }

def main():
    """Hàm test và demo"""
    
    # Khởi tạo processor
    processor = ImageNamingProcessor(
        domain="https://myshop.com/product/",
        image_base="https://cdn.myshop.com/images/"
    )
    
    # Test với một số mã sản phẩm
    test_codes = [
        "FR-1H-220V",
        "TC-NT-20R (with special coating)",
        "ETC-48 Add-on Kit",
        "FR-2H-380V [with special coating]",
        "TC-NT-30R ADD ON KIT",
        "ETC-60",
        "FR-3H-440V"
    ]
    
    print("🧪 TEST XỬ LÝ MÃ SẢN PHẨM")
    print("=" * 50)
    
    for code in test_codes:
        result = processor.process_single_code(code)
        print(f"\n📋 Mã gốc: {result['original_code']}")
        print(f"🧹 Mã sạch: {result['clean_code']}")
        print(f"🖼️ Tên ảnh: {result['image_name']}")
        print(f"📦 Add-on kit: {result['had_addon_kit']}")
        print(f"🔗 Product URL: {result['product_url']}")
        print(f"🖼️ Image URL: {result['image_url']}")
        print(f"📄 Filename: {result['filename']}")
    
    print("\n" + "=" * 50)
    print("📝 Cách sử dụng:")
    print("1. processor = ImageNamingProcessor()")
    print("2. result = processor.process_single_code('FR-1H-220V')")
    print("3. filename = result['filename']")
    print("\nHoặc xử lý file Excel:")
    print("processor.process_excel_file('input.xlsx', 'output.xlsx')")

if __name__ == "__main__":
    main()
