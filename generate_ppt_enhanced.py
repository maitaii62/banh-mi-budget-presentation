from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.oxml.xmlchemy import OxmlElement

OUTPUT_FILE = "Ke_hoach_ngan_sach_quan_banh_mi_buoi_sang_enhanced.pptx"

# Color palette
COLOR_PRIMARY = RGBColor(23, 62, 108)      # Deep blue
COLOR_SECONDARY = RGBColor(70, 130, 180)  # Steel blue
COLOR_ACCENT = RGBColor(255, 140, 0)      # Orange accent
COLOR_TEXT = RGBColor(40, 40, 40)         # Dark gray
COLOR_LIGHT_BG = RGBColor(248, 250, 252)  # Very light blue
COLOR_WHITE = RGBColor(255, 255, 255)
COLOR_SUCCESS = RGBColor(34, 139, 34)     # Forest green

def add_animation_fade(shape, delay=0):
    """Add fade-in animation to shape"""
    sp = shape.element
    timing = OxmlElement('p:timing')
    tn_lst = OxmlElement('p:tnLst')
    par = OxmlElement('p:par')
    
    cTn = OxmlElement('p:cTn')
    cTn.set('id', '1')
    cTn.set('dur', '500')
    cTn.set('restart', 'never')
    
    par.append(cTn)
    tn_lst.append(par)
    timing.append(tn_lst)
    
    if sp.find('.//{http://schemas.openxmlformats.org/presentationml/2006/main}timing') is None:
        sp.append(timing)

def add_gradient_fill(shape, color1, color2):
    """Add gradient fill to shape"""
    fill = shape.fill
    fill.gradient()
    fill.gradient_angle = 90.0
    fill.gradient_stops[0].color.rgb = color1
    fill.gradient_stops[1].color.rgb = color2

def set_background_gradient(slide, color1, color2):
    """Set gradient background for slide"""
    fill = slide.background.fill
    fill.gradient()
    fill.gradient_angle = 45.0
    fill.gradient_stops[0].color.rgb = color1
    fill.gradient_stops[1].color.rgb = color2

def add_title_with_line(slide, title_text):
    """Add title with decorative line"""
    # Title
    title_box = slide.shapes.title
    title_box.text = title_text
    title_box.text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT
    title_box.text_frame.paragraphs[0].font.size = Pt(32)
    title_box.text_frame.paragraphs[0].font.bold = True
    title_box.text_frame.paragraphs[0].font.color.rgb = COLOR_PRIMARY
    
    # Decorative line under title
    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(0.8), Inches(0.95), Inches(3.0), Inches(0.08)
    )
    line.fill.solid()
    line.fill.fore_color.rgb = COLOR_ACCENT
    line.line.color.rgb = COLOR_ACCENT

def add_decorated_bullet_box(slide, bullets, left=0.8, top=1.6, width=11.5, height=4.8):
    """Add bullets with decorative background box"""
    # Background box
    bg_box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(left - 0.1), Inches(top - 0.1), Inches(width + 0.2), Inches(height + 0.2)
    )
    bg_box.fill.solid()
    bg_box.fill.fore_color.rgb = RGBColor(245, 248, 252)
    bg_box.line.color.rgb = RGBColor(200, 220, 240)
    bg_box.line.width = Pt(1)
    bg_box.shadow.inherit = False
    
    # Text box with bullets
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP
    
    for idx, item in enumerate(bullets):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.text = item
        p.level = 0
        p.bullet = True
        p.font.name = "Calibri"
        p.font.size = Pt(18)
        p.font.color.rgb = COLOR_TEXT
        p.space_after = Pt(10)
        p.space_before = Pt(4)

def add_stat_card(slide, label, value, unit="", left=0.8, top=1.5, width=3.2, height=1.8):
    """Add beautiful stat card"""
    # Card background with gradient
    card = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    add_gradient_fill(card, RGBColor(230, 245, 255), RGBColor(200, 230, 250))
    card.line.color.rgb = COLOR_SECONDARY
    card.line.width = Pt(2)
    
    # Add shadow
    card.shadow.inherit = False
    card.shadow.shadow_type = 'outer'
    
    text_frame = card.text_frame
    text_frame.clear()
    text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    text_frame.margin_bottom = Inches(0.05)
    text_frame.margin_top = Inches(0.05)
    text_frame.margin_left = Inches(0.08)
    text_frame.margin_right = Inches(0.08)
    
    # Label
    p1 = text_frame.paragraphs[0]
    p1.text = label
    p1.alignment = PP_ALIGN.CENTER
    p1.font.size = Pt(12)
    p1.font.bold = True
    p1.font.color.rgb = COLOR_SECONDARY
    
    # Value
    p2 = text_frame.add_paragraph()
    p2.text = value
    p2.alignment = PP_ALIGN.CENTER
    p2.font.size = Pt(24)
    p2.font.bold = True
    p2.font.color.rgb = COLOR_SUCCESS
    
    # Unit
    if unit:
        p3 = text_frame.add_paragraph()
        p3.text = unit
        p3.alignment = PP_ALIGN.CENTER
        p3.font.size = Pt(11)
        p3.font.color.rgb = RGBColor(100, 100, 100)

def add_enhanced_table(slide, rows, title="", left=0.6, top=1.5, width=12.1, height=4.5):
    """Add table with enhanced styling"""
    if title:
        title_box = slide.shapes.add_textbox(Inches(left), Inches(top - 0.5), Inches(8), Inches(0.4))
        tf = title_box.text_frame
        p = tf.paragraphs[0]
        p.text = title
        p.font.name = "Calibri"
        p.font.size = Pt(18)
        p.font.bold = True
        p.font.color.rgb = COLOR_PRIMARY
    
    table_shape = slide.shapes.add_table(len(rows), len(rows[0]), Inches(left), Inches(top), Inches(width), Inches(height))
    table = table_shape.table
    
    # Set column widths
    if len(rows[0]) == 2:
        table.columns[0].width = Inches(6)
        table.columns[1].width = Inches(6.1)
    elif len(rows[0]) == 3:
        for col in table.columns:
            col.width = Inches(4)
    elif len(rows[0]) == 4:
        table.columns[0].width = Inches(3.5)
        table.columns[1].width = Inches(2.3)
        table.columns[2].width = Inches(2.8)
        table.columns[3].width = Inches(3.5)
    
    for i, row in enumerate(rows):
        for j, value in enumerate(row):
            cell = table.cell(i, j)
            cell.text = value
            
            # Header row styling
            if i == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = COLOR_PRIMARY
                for paragraph in cell.text_frame.paragraphs:
                    paragraph.alignment = PP_ALIGN.CENTER
                    paragraph.font.name = "Calibri"
                    paragraph.font.size = Pt(12)
                    paragraph.font.bold = True
                    paragraph.font.color.rgb = COLOR_WHITE
                    paragraph.space_after = Pt(6)
            # Data rows
            else:
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(248, 250, 252) if i % 2 == 0 else RGBColor(240, 245, 250)
                for paragraph in cell.text_frame.paragraphs:
                    paragraph.alignment = PP_ALIGN.CENTER if i == 0 else PP_ALIGN.LEFT
                    paragraph.font.name = "Calibri"
                    paragraph.font.size = Pt(11)
                    paragraph.font.color.rgb = COLOR_TEXT

def create_presentation():
    """Create enhanced presentation"""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    
    # Slide 1 - Cover
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_background_gradient(slide, COLOR_PRIMARY, COLOR_SECONDARY)
    
    # Logo area
    logo = slide.shapes.add_textbox(Inches(0.8), Inches(0.6), Inches(3.5), Inches(0.6))
    logo_tf = logo.text_frame
    logo_tf.text = "🥖 DỰ ÁN KINH DOANH"
    logo_tf.paragraphs[0].font.name = "Calibri"
    logo_tf.paragraphs[0].font.bold = True
    logo_tf.paragraphs[0].font.size = Pt(16)
    logo_tf.paragraphs[0].font.color.rgb = COLOR_WHITE
    
    # Main title
    cover_title = slide.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(11.6), Inches(2.0))
    cover_title_tf = cover_title.text_frame
    cover_title_tf.word_wrap = True
    
    p1 = cover_title_tf.paragraphs[0]
    p1.text = "KẾ HOẠCH NGÂN SÁCH"
    p1.font.name = "Calibri"
    p1.font.bold = True
    p1.font.size = Pt(44)
    p1.font.color.rgb = COLOR_WHITE
    
    p2 = cover_title_tf.add_paragraph()
    p2.text = "Dự án quán bánh mì buổi sáng"
    p2.font.name = "Calibri"
    p2.font.size = Pt(28)
    p2.font.color.rgb = RGBColor(200, 220, 240)
    
    # Info cards on cover
    info_cards = [
        ("Sản lượng mục tiêu", "50 ổ/ngày", Inches(0.8)),
        ("Doanh thu dự kiến", "32,5 triệu/tháng", Inches(4.8)),
        ("Chi phí trước dự phòng", "24,0 triệu/tháng", Inches(8.8))
    ]
    
    for label, value, left_pos in info_cards:
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left_pos, Inches(4.3), Inches(3.8), Inches(1.3))
        card.fill.solid()
        card.fill.fore_color.rgb = RGBColor(240, 250, 255)
        card.line.color.rgb = COLOR_ACCENT
        card.line.width = Pt(2)
        
        tf = card.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.word_wrap = True
        
        p_label = tf.paragraphs[0]
        p_label.text = label
        p_label.font.name = "Calibri"
        p_label.font.size = Pt(12)
        p_label.font.bold = True
        p_label.font.color.rgb = COLOR_SECONDARY
        p_label.alignment = PP_ALIGN.CENTER
        
        p_value = tf.add_paragraph()
        p_value.text = value
        p_value.font.name = "Calibri"
        p_value.font.size = Pt(16)
        p_value.font.bold = True
        p_value.font.color.rgb = COLOR_ACCENT
        p_value.alignment = PP_ALIGN.CENTER
    
    # Slide 2 - Mô tả dự án
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_background_gradient(slide, COLOR_LIGHT_BG, RGBColor(255, 255, 255))
    add_title_with_line(slide, "1. MÔ TẢ DỰ ÁN")
    
    bullets = [
        "Tên dự án: Bánh mì buổi sáng – Ăn nhanh, ngon và tiện lợi.",
        "Sản phẩm: bánh mì thịt, bánh mì trứng, bánh mì đặc biệt.",
        "Khách hàng mục tiêu: học sinh, sinh viên, nhân viên văn phòng, người dân quanh khu vực.",
        "Hình thức kinh doanh: bán trực tiếp tại quầy, nhận đơn qua điện thoại và mạng xã hội.",
        "Quy mô: 1 quầy bánh mì, hoạt động chủ yếu vào buổi sáng, 50 ổ/ngày."
    ]
    add_decorated_bullet_box(slide, bullets, left=0.7, top=1.6, width=11.6, height=4.5)
    
    # Slide 3 - Mục tiêu ngân sách
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_background_gradient(slide, COLOR_LIGHT_BG, RGBColor(255, 255, 255))
    add_title_with_line(slide, "2. MỤC TIÊU NGÂN SÁCH")
    
    add_stat_card(slide, "Doanh thu dự kiến", "32,5 triệu", "₫/tháng", left=0.7, top=1.7)
    add_stat_card(slide, "Chi phí cố định", "4,5 triệu", "₫/tháng", left=4.3, top=1.7)
    add_stat_card(slide, "Chi phí biến đổi", "19,5 triệu", "₫/tháng", left=7.9, top=1.7)
    add_stat_card(slide, "Dự phòng 5%", "1,2 triệu", "₫/tháng", left=11.0, top=1.7)
    
    bullets2 = [
        "Mục tiêu doanh thu: 32,5 triệu/tháng từ 50 ổ/ngày.",
        "Mục tiêu kiểm soát chi phí: giữ chi phí trong ngân sách được phê duyệt.",
        "Mục tiêu lợi nhuận: đảm bảo chênh lệch thu – chi sau dự phòng."
    ]
    add_decorated_bullet_box(slide, bullets2, left=0.7, top=4.1, width=11.6, height=2.0)
    
    # Slide 4 - Doanh thu dự kiến
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_background_gradient(slide, COLOR_LIGHT_BG, RGBColor(255, 255, 255))
    add_title_with_line(slide, "3. DOANH THU DỰ KIẾN")
    
    rows = [
        ["Sản phẩm", "Số lượng/tháng", "Giá bán TB", "Doanh thu/tháng"],
        ["Bánh mì các loại", "1.300 ổ", "25.000₫/ổ", "32.500.000₫"],
        ["Tổng cộng", "", "", "32.500.000₫"]
    ]
    add_enhanced_table(slide, rows, title="Công thức: 50 ổ/ngày × 26 ngày = 1.300 ổ/tháng")
    
    # Slide 5 - Chi phí dự kiến
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_background_gradient(slide, COLOR_LIGHT_BG, RGBColor(255, 255, 255))
    add_title_with_line(slide, "4. CHI PHÍ DỰ KIẾN")
    
    rows1 = [
        ["Chi phí cố định", "Chi phí/tháng"],
        ["Thuê địa điểm", "3.000.000₫"],
        ["Điện, nước", "500.000₫"],
        ["Quảng cáo, in tờ rơi", "500.000₫"],
        ["Khấu hao thiết bị", "500.000₫"],
        ["TỔNG", "4.500.000₫"]
    ]
    
    rows2 = [
        ["Chi phí biến đổi", "Chi phí/tháng"],
        ["Nguyên liệu làm bánh", "16.900.000₫"],
        ["Bao bì, túi giấy", "1.300.000₫"],
        ["Gia vị, rau, vật liệu phụ", "1.300.000₫"],
        ["TỔNG", "19.500.000₫"]
    ]
    
    add_enhanced_table(slide, rows1, left=0.6, top=1.5, width=5.9, height=3.5)
    add_enhanced_table(slide, rows2, left=7.0, top=1.5, width=5.7, height=3.5)
    
    # Slide 6 - Bảng tổng hợp
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_background_gradient(slide, COLOR_LIGHT_BG, RGBColor(255, 255, 255))
    add_title_with_line(slide, "5. BẢNG TỔNG HỢP NGÂN SÁCH")
    
    rows = [
        ["Chỉ tiêu", "Giá trị"],
        ["Sản lượng dự kiến", "1.300 ổ/tháng"],
        ["Giá bán trung bình", "25.000₫/ổ"],
        ["Doanh thu dự kiến", "32.500.000₫"],
        ["Chi phí cố định", "4.500.000₫"],
        ["Chi phí biến đổi", "19.500.000₫"],
        ["Tổng chi phí trước dự phòng", "24.000.000₫"],
        ["Dự phòng 5%", "1.200.000₫"],
        ["Tổng ngân sách", "25.200.000₫"],
        ["Lợi nhuận ước tính", "7.300.000₫"]
    ]
    add_enhanced_table(slide, rows, left=1.5, top=1.5, width=10.3, height=5.0)
    
    # Slide 7 - Lợi nhuận dự kiến
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_background_gradient(slide, COLOR_LIGHT_BG, RGBColor(255, 255, 255))
    add_title_with_line(slide, "6. LỢI NHUẬN DỰ KIẾN")
    
    add_stat_card(slide, "Chênh lệch trước dự phòng", "8,5 triệu", "₫", left=0.9, top=1.7)
    add_stat_card(slide, "Sau khi trích dự phòng", "7,3 triệu", "₫", left=4.5, top=1.7)
    add_stat_card(slide, "Tỷ lệ sinh lợi", "22,46%", "trên doanh thu", left=8.1, top=1.7)
    
    bullets3 = [
        "Doanh thu 32,5 triệu – Chi phí 24,0 triệu = Chênh lệch 8,5 triệu.",
        "Sau khi trích dự phòng 1,2 triệu, còn lại 7,3 triệu/tháng.",
        "Tỷ lệ sinh lợi: 7,3 ÷ 32,5 × 100% ≈ 22,46% doanh thu."
    ]
    add_decorated_bullet_box(slide, bullets3, left=0.8, top=4.2, width=11.6, height=2.0)
    
    # Slide 8 - Nguồn tài trợ
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_background_gradient(slide, COLOR_LIGHT_BG, RGBColor(255, 255, 255))
    add_title_with_line(slide, "7. NGUỒN TÀI TRỢ")
    
    rows = [
        ["Nguồn vốn", "Số tiền"],
        ["Vốn tự có", "8.000.000₫"],
        ["Vốn góp của các thành viên", "5.000.000₫"],
        ["Vay người thân", "5.000.000₫"],
        ["TỔNG CỘNG", "18.000.000₫"]
    ]
    add_enhanced_table(slide, rows, left=2.0, top=1.5, width=9.3, height=3.2)
    
    bullets4 = [
        "Vốn ban đầu dùng để: chuẩn bị quầy hàng, mua dụng cụ, đặt cọc địa điểm, mua nguyên liệu ban đầu.",
        "Lưu ý: Vốn 18 triệu < ngân sách 25,2 triệu/tháng, nên cần bổ sung từ doanh thu hoặc nguồn hợp pháp khác."
    ]
    add_decorated_bullet_box(slide, bullets4, left=0.8, top=5.0, width=11.6, height=1.8)
    
    # Slide 9 - Thời gian thực hiện
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_background_gradient(slide, COLOR_LIGHT_BG, RGBColor(255, 255, 255))
    add_title_with_line(slide, "8. THỜI GIAN THỰC HIỆN (8 TUẦN)")
    
    rows = [
        ["Tuần", "Công việc chính"],
        ["Tuần 1", "Khảo sát địa điểm, nhu cầu, đối thủ cạnh tranh"],
        ["Tuần 2", "Tìm nhà cung cấp nguyên liệu, chuẩn bị dụng cụ"],
        ["Tuần 3", "Chuẩn bị quầy hàng, thiết kế bảng hiệu"],
        ["Tuần 4", "Mua nguyên liệu, bán thử, điều chỉnh quy trình"],
        ["Tuần 5-7", "Kinh doanh chính thức, theo dõi doanh thu, chi phí"],
        ["Tuần 8", "Tổng kết, đánh giá, rút kinh nghiệm"]
    ]
    add_enhanced_table(slide, rows, left=0.6, top=1.5, width=12.1, height=5.0)
    
    # Slide 10 - Quy trình đánh giá kết quả
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_background_gradient(slide, COLOR_LIGHT_BG, RGBColor(255, 255, 255))
    add_title_with_line(slide, "9. QUY TRÌNH ĐÁNH GIÁ KẾT QUẢ NGÂN SÁCH")
    
    bullets5 = [
        "Bước 1: Đánh giá chi tiêu thực tế so với kế hoạch.",
        "Bước 2: Đánh giá doanh thu và chi phí thực tế.",
        "Bước 3: Xem xét mục tiêu và chỉ tiêu ngân sách.",
        "Bước 4: Điều chỉnh ngân sách phù hợp với thực tế.",
        "Bước 5: Xác định và khắc phục thất thoát ngân sách.",
        "Bước 6: Rà soát định kỳ và rút kinh nghiệm."
    ]
    add_decorated_bullet_box(slide, bullets5, left=0.8, top=1.6, width=11.6, height=4.7)
    
    # Slide 11 - Điểm hòa vốn
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_background_gradient(slide, COLOR_LIGHT_BG, RGBColor(255, 255, 255))
    add_title_with_line(slide, "10. ĐIỂM HÒA VỐN")
    
    add_stat_card(slide, "Chi phí cố định", "4,5 triệu", "₫/tháng", left=1.0, top=1.7)
    add_stat_card(slide, "Giá bán TB - Chi phí biến TB", "10.000₫", "/ổ", left=5.0, top=1.7)
    add_stat_card(slide, "Điểm hòa vốn", "450 ổ", "/tháng (≈18 ổ/ngày)", left=8.5, top=1.7)
    
    bullets6 = [
        "Công thức: Chi phí cố định ÷ (Giá bán TB – Chi phí biến TB) = 4.500.000 ÷ 10.000 = 450 ổ.",
        "Tính bình quân ngày: 450 ÷ 26 ≈ 17,31 ≈ 18 ổ/ngày.",
        "Kết luận: Quán cần bán tối thiểu 18 ổ/ngày để không lỗ."
    ]
    add_decorated_bullet_box(slide, bullets6, left=0.8, top=4.1, width=11.6, height=2.0)
    
    # Slide 12 - Phương án cắt giảm 15%
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_background_gradient(slide, COLOR_LIGHT_BG, RGBColor(255, 255, 255))
    add_title_with_line(slide, "11. PHƯƠNG ÁN CẮT GIẢM 15% CHI PHÍ")
    
    add_stat_card(slide, "Chi phí cũ", "24,0 triệu", "₫/tháng", left=1.2, top=1.7)
    add_stat_card(slide, "Mức tiết kiệm", "3,6 triệu", "15%", left=4.8, top=1.7)
    add_stat_card(slide, "Chi phí mới", "20,4 triệu", "₫/tháng", left=8.4, top=1.7)
    
    bullets7 = [
        "Mục tiêu: Giảm chi phí 3,6 triệu/tháng (15% tổng chi phí) mà vẫn bảo đảm chất lượng.",
        "Ưu tiên cắt giảm: nguyên liệu (2,5M), gia vị (0,4M), quảng cáo (0,3M), điện nước (0,2M), bao bì (0,2M).",
        "Không cắt: Thuê địa điểm và khấu hao thiết bị (chi phí cố định, khó điều chỉnh)."
    ]
    add_decorated_bullet_box(slide, bullets7, left=0.8, top=4.1, width=11.6, height=2.0)
    
    # Slide 13 - Chi tiết cắt giảm
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_background_gradient(slide, COLOR_LIGHT_BG, RGBColor(255, 255, 255))
    add_title_with_line(slide, "12. BẢNG CẮT GIẢM CHI TIẾT")
    
    rows = [
        ["Khoản mục", "Ban đầu", "Cắt giảm", "Sau điều chỉnh"],
        ["Nguyên liệu làm bánh", "16,9M", "2,5M", "14,4M"],
        ["Bao bì, túi giấy", "1,3M", "0,2M", "1,1M"],
        ["Gia vị, rau, vật liệu phụ", "1,3M", "0,4M", "0,9M"],
        ["Quảng cáo, in tờ rơi", "0,5M", "0,3M", "0,2M"],
        ["Điện, nước", "0,5M", "0,2M", "0,3M"],
        ["Thuê địa điểm", "3,0M", "0M", "3,0M"],
        ["Khấu hao thiết bị", "0,5M", "0M", "0,5M"],
        ["TỔNG CỘNG", "24,0M", "3,6M", "20,4M"]
    ]
    add_enhanced_table(slide, rows, left=0.6, top=1.5, width=12.1, height=4.8)
    
    # Slide 14 - Kế hoạch dự phòng
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_background_gradient(slide, COLOR_LIGHT_BG, RGBColor(255, 255, 255))
    add_title_with_line(slide, "13. KẾ HOẠCH DỰ PHÒNG")
    
    bullets8 = [
        "Có nên lập dự phòng không? CÓ – để chủ động ứng phó với giá nguyên liệu tăng, doanh thu giảm, thiết bị hư hỏng.",
        "Các bước: Xác định rủi ro → Ước tính chi phí → Chọn tỷ lệ dự phòng → Xây dựng quy định sử dụng → Theo dõi điều chỉnh.",
        "Tỷ lệ dự phòng: 5% × 24.000.000₫ = 1.200.000₫/tháng.",
        "Tổng ngân sách bao gồm dự phòng: 24.000.000 + 1.200.000 = 25.200.000₫/tháng.",
        "Nguyên tắc: Dự phòng chỉ dùng khi cần thiết, phải ghi nhận rõ nguyên nhân, không tự động tính chi phí."
    ]
    add_decorated_bullet_box(slide, bullets8, left=0.8, top=1.6, width=11.6, height=4.7)
    
    # Slide 15 - Kiểm soát ngân sách
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_background_gradient(slide, COLOR_LIGHT_BG, RGBColor(255, 255, 255))
    add_title_with_line(slide, "14. KIỂM SOÁT NGÂN SÁCH")
    
    bullets9 = [
        "✓ Ghi nhận doanh thu & chi phí hàng ngày để theo dõi tình hình kinh doanh.",
        "✓ Kiểm kê nguyên liệu cuối ngày, đối chiếu nhập/xuất/tồn kho phát hiện hao hụt.",
        "✓ So sánh giá từ nhiều nhà cung cấp trước khi mua hàng.",
        "✓ Đối chiếu chi phí thực tế với ngân sách hằng tuần, phát hiện chênh lệch.",
        "✓ Duy trì quỹ dự phòng 1.200.000₫/tháng, ghi nhận việc sử dụng rõ ràng.",
        "✓ Khi vượt kế hoạch: xác định nguyên nhân (giá tăng, sản lượng giảm, hao hụt...) và điều chỉnh kịp thời."
    ]
    add_decorated_bullet_box(slide, bullets9, left=0.8, top=1.6, width=11.6, height=4.7)
    
    # Slide 16 - Kết luận
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_background_gradient(slide, COLOR_PRIMARY, COLOR_SECONDARY)
    
    # Title on dark background
    title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.8), Inches(11.6), Inches(0.8))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = "15. KẾT LUẬN"
    p.font.name = "Calibri"
    p.font.size = Pt(32)
    p.font.bold = True
    p.font.color.rgb = COLOR_WHITE
    
    # Decorative line
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.65), Inches(3.0), Inches(0.06))
    line.fill.solid()
    line.fill.fore_color.rgb = COLOR_ACCENT
    line.line.color.rgb = COLOR_ACCENT
    
    bullets10 = [
        "✓ Dự án quán bánh mì có tính khả thi cao với quy mô nhỏ, địa điểm tập trung.",
        "✓ Nếu đạt doanh thu 32,5 triệu/tháng & kiểm soát chi phí tốt, sinh lợi 7,3 triệu/tháng (22,46%).",
        "✓ Cắt giảm 15% chi phí (3,6 triệu) hoàn toàn khả thi thông qua quản lý nguyên liệu & kiểm soát chi phí.",
        "✓ Dự phòng 5% (1,2 triệu) giúp chủ động ứng phó với rủi ro & biến động thị trường.",
        "✓ Thực hiện thành công cần: theo dõi doanh thu hàng ngày, kiểm kê định kỳ, điều chỉnh kế hoạch kịp thời."
    ]
    
    box = slide.shapes.add_textbox(Inches(0.8), Inches(2.0), Inches(11.6), Inches(4.8))
    tf = box.text_frame
    tf.word_wrap = True
    
    for idx, item in enumerate(bullets10):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.text = item
        p.level = 0
        p.font.name = "Calibri"
        p.font.size = Pt(18)
        p.font.color.rgb = COLOR_WHITE
        p.space_after = Pt(12)
    
    # Save presentation
    prs.save(OUTPUT_FILE)
    print(f"✓ Enhanced PowerPoint created: {OUTPUT_FILE}")

if __name__ == "__main__":
    create_presentation()
