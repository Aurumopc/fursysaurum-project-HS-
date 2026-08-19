import pandas as pd
import json
import os
import re

def clean_prefix(prefix_str):
    """품번 Prefix 내 특수문자 및 기호 정리"""
    if pd.isna(prefix_str):
        return "UNKNOWN"
    cleaned = str(prefix_str).strip()
    cleaned_code = re.sub(r'[^a-zA-Z0-9]', '_', cleaned)
    cleaned_code = re.sub(r'_+', '_', cleaned_code).strip('_')
    return cleaned_code if cleaned_code else "FURSYS"

def map_category(row):
    """1. 대분류를 한글 쇼핑몰 카테고리로 변환"""
    big_cat = str(row['대분류']).strip()
    p_type = str(row['제품유형 (쇼핑몰 상품명 단위)']).strip()
    
    if big_cat == 'Seating':
        return '의자'
    elif big_cat == 'Tables':
        return '테이블'
    elif big_cat == 'Sofas & Armchairs':
        return '소파/라운지'
    elif big_cat == 'Space in Space':
        return '파티션'
    elif big_cat == 'Common Items':
        if any(keyword in p_type for keyword in ['캐비닛', '선반장', '라커', '옷장', '스토리지']):
            return '수납장'
        elif any(keyword in p_type for keyword in ['패널', '디바이더']):
            return '파티션'
        else:
            return '수납장'
    elif big_cat == 'Workspaces':
        if any(keyword in p_type for keyword in ['스크린', '패널']):
            return '파티션'
        elif any(keyword in p_type for keyword in ['스토리지', '옷장', '서랍', '캐비닛', '페그보드', '라커']):
            return '수납장'
        else:
            return '책상'
    return '책상'

def map_space(category, p_type, series):
    """공간별 카테고리 태그 자동 부여"""
    p_str = f"{p_type} {series}"
    if any(k in p_str for k in ['중역', '임원', '프리미엄']):
        return '임원공간'
    elif category == '테이블' or '회의' in p_str:
        return '회의공간'
    elif category == '소파/라운지' or any(k in p_str for k in ['소파', '휴게', '라운지', '스툴']):
        return '라운지/휴게공간'
    else:
        return '업무공간'

def convert_excel_to_json(excel_path, output_json_path):
    # 엑셀 파일 로드
    df_prod = pd.read_excel(excel_path, sheet_name='제품목록')
    df_opt = pd.read_excel(excel_path, sheet_name='옵션그룹상세')
    
    # 옵션 데이터를 (시리즈, 제품유형) 키로 딕셔너리화
    opt_dict = {}
    for _, opt_row in df_opt.iterrows():
        key = (str(opt_row['시리즈']).strip(), str(opt_row['제품유형']).strip())
        if key not in opt_dict:
            opt_dict[key] = []
        
        opt_group = str(opt_row['옵션그룹']).strip()
        opt_vals_raw = str(opt_row['옵션값 (등록용 리스트)']).strip()
        opt_vals = [v.strip() for v in opt_vals_raw.split('/') if v.strip()]
        
        opt_dict[key].append({
            "name": opt_group,
            "values": opt_vals
        })
        
    products = []
    for idx, row in df_prod.iterrows():
        series = str(row['시리즈']).strip()
        p_type = str(row['제품유형 (쇼핑몰 상품명 단위)']).strip()
        raw_prefix = str(row['품번 Prefix']).strip()
        code = clean_prefix(raw_prefix)  # 2. 품번 Prefix 기반 Key 설정
        
        category = map_category(row)      # 1. 한글 카테고리 변환
        space = map_space(category, p_type, series)
        
        # 옵션 그룹 매칭
        key = (series, p_type)
        product_options = []
        if key in opt_dict:
            product_options = opt_dict[key]
        else:
            size_opt = str(row['사이즈 옵션']).strip() if pd.notna(row['사이즈 옵션']) else ""
            if size_opt and size_opt != "nan":
                sizes = [s.strip() for s in size_opt.split('/') if s.strip()]
                product_options.append({"name": "사이즈", "values": sizes})

        # 3. 이미지 URL 객체 생성 규칙 반영
        product = {
            "no": int(row['No']),
            "product_code": code,
            "raw_prefix": raw_prefix,
            "product_name": p_type,
            "series": series,
            "category": category,
            "space_category": space,
            "base_price": 350000 + (idx % 8) * 40000,
            "options": product_options,
            "images": {
                "main": f"https://raw.githubusercontent.com/[GITHUB_ID]/fursys-assets/main/images/{code}/main.jpg",
                "pip_sub": f"https://raw.githubusercontent.com/[GITHUB_ID]/fursys-assets/main/images/{code}/sub_pip.jpg"
            }
        }
        products.append(product)
        
    # 4. fursys-store/data/products.json 으로 저장
    os.makedirs(os.path.dirname(output_json_path), exist_ok=True)
    with open(output_json_path, 'w', encoding='utf-8') as f:
        json.dump(products, f, ensure_ascii=False, indent=2)
        
    print(f"총 {len(products)}개 제품 변환 완료 -> {output_json_path}")

if __name__ == '__main__':
    # 엑셀파일명에 맞춰 실행
    excel_file = 'fursys_data.xlsx'
    if not os.path.exists(excel_file) and os.path.exists('fursys_data.xlsx.xlsx'):
        excel_file = 'fursys_data.xlsx.xlsx'
        
    convert_excel_to_json(excel_file, 'fursys-store/data/products.json')