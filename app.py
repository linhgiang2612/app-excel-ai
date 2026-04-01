import streamlit as st
import pandas as pd
import re # Thư viện để xử lý chuỗi nâng cao (kiểm tra công thức)

# --- CẤU HÌNH TRANG ---
st.set_page_config(page_title="Học Excel Cơ Bản", layout="wide")

# --- HÀM CHO PHẦN LÝ THUYẾT ---

def hien_thi_ly_thuyet():
    st.header("📘 CHƯƠNG 3: SỬ DỤNG BẢNG TÍNH CƠ BẢN")
    st.markdown("""
    **Mục tiêu:** Hiểu ý nghĩa, cú pháp, cách vận dụng các hàm thông dụng và kỹ thuật trình bày dữ liệu.
    """)
    st.divider()

    # --- MỤC 3.1: GIỚI THIỆU ---
    with st.expander("3.1. Giới thiệu Giao diện & Kiểu dữ liệu", expanded=True):
        st.subheader("3.1.1. Giao diện màn hình chính")
        st.markdown("""
        Các thành phần cơ bản trên màn hình Excel:
        * **(1) Thanh công cụ nhanh**: Chứa các lệnh thao tác nhanh.
        * **(3) Thanh công cụ Ribbon**: Chứa các lệnh thao tác được phân chia thành các nhóm.
        * **(4) Name box**: Vùng địa chỉ vị trí con trỏ hiện thời.
        * **(5) Formula bar**: Thanh công thức.
        * **(6) Màn hình nhập liệu (WorkSheet)**: Vùng lớn nhất chứa dữ liệu bảng tính.
        * **(9) Thanh Sheet tab**: Liệt kê danh sách các bảng tính (Worksheet) có trong tệp (Workbook).
        """)

        st.subheader("3.1.2. Kiểu dữ liệu và Phương pháp nhập")
        st.markdown("""
        **Các kiểu dữ liệu chính:**
        * **Dữ liệu kiểu số**: Gồm số 0-9 và ký tự đặc biệt ($+; -; \$; \%$). Mặc định căn lề **Phải**. 
            * *Lưu ý:* Nếu xuất hiện dãy `###`, cần kéo rộng cột để hiển thị.
        * **Dữ liệu kiểu chuỗi**: Ký tự đầu là chữ hoặc số. Mặc định căn lề **Trái**. 
            * *Lưu ý:* Để nhập số dạng chuỗi, gõ dấu nháy đơn (`'`) trước số.
        * **Dữ liệu kiểu ngày tháng**: Là một dạng của dữ liệu số. Thiết lập định dạng tại Control Panel theo dạng $dd/mm/yyyy$.
        * **Dữ liệu kiểu logic**: Gồm hai giá trị **TRUE** (Đúng) và **FALSE** (Sai).
        
        **Nguyên tắc nhập liệu:**
        * **Dữ liệu cố định**: Chọn ô -> Nhập dữ liệu -> Nhấn Enter.
        * **Dữ liệu công thức**: Bắt đầu bằng dấu `=` hoặc `+`.
        """)

    # --- MỤC 3.2: QUẢN LÝ BẢNG TÍNH ---
    with st.expander("3.2. Quản lý Bảng tính & Địa chỉ"):
        st.subheader("3.2.1. Các khái niệm về Ô và Vùng")
        st.markdown("""
        * **Ô (Cell)**: Giao của cột và dòng. Địa chỉ có cấu trúc: `<CỘT><DÒNG>` (VD: $A1, AA3$).
        * **Vùng (Range)**: Tập hợp nhiều ô. Địa chỉ: `<Ô TRÊN TRÁI>:<Ô DƯỚI PHẢI>` (VD: $A1:C5$).
        """)

        st.subheader("3.2.3. Các loại địa chỉ")
        st.markdown("""
        * **Địa chỉ tương đối**: Thay đổi khi sao chép công thức. Dạng: `A1, B3`.
        * **Địa chỉ tuyệt đối**: Không thay đổi khi sao chép công thức. Dạng: `\$A\$1, \$B\$3`.
        * **Địa chỉ hỗn hợp**: Kết hợp giữa tương đối và tuyệt đối. Dạng: `\$B2` (cố định cột) hoặc `B\$2` (cố định dòng).
        """)

        st.subheader("Thông báo lỗi thường gặp")
        st.table({
            "Mã lỗi": ["#DIV/0!", "#NAME?", "#VALUE!", "#REF!", "#N/A"],
            "Lý do": ["Chia cho số 0", "Sai tên hàm/tham chiếu", "Kiểu dữ liệu tính toán sai", "Vùng tham chiếu sai", "Dữ liệu không tồn tại (hàm dò tìm)"]
        }) # [cite: 580]

    # --- MỤC 3.3 ĐỊNH DẠNG ---
    with st.expander("3.3 Định dạng"):
        st.subheader("3.3. Định dạng bảng tính")
        st.markdown("""
        Sử dụng hộp thoại **Format Cells (Ctrl + 1)**:
        * **Alignment**: Định dạng vị trí văn bản (ngang, dọc, xoay chữ).
            * *Wrap text*: Xuống hàng trong cùng 1 ô.
            * *Merge cells*: Nối nhiều ô thành một.
        * **Font**: Định dạng kiểu chữ, kích cỡ, màu sắc và hiệu ứng (chỉ số trên/dưới).
        * **Number**: Định dạng hiển thị số (Tiền tệ, Ngày tháng, Phần trăm...).
        """)

    # --- MỤC 3.5: CÔNG THỨC VÀ HÀM ---
    st.subheader("3.5. Thao tác với Công thức và Hàm")
    
    st.markdown("#### 3.5.1. Toán tử ")
    st.markdown("""
    * **Số học**: `+`, `-`, `*`, `/`, `%`, `^` (lũy thừa).
    * **So sánh**: `>`, `<`, `=`, `<>`, `>=`, `<=`.
    * **Nối chuỗi**: `&`.
    """)

# --- 3.5.3: HỆ THỐNG CÁC HÀM ---
    st.subheader("3.5.3. Các hàm cơ bản thường dùng")

    # --- NHÓM HÀM THỐNG KÊ ---
    with st.container():
        st.info("📊 **NHÓM HÀM THỐNG KÊ**")
        
        # Hàm SUM
        st.markdown("#### 1. Hàm SUM")
        st.latex(r"=SUM(X_1, X_2, ..., X_{255}) \text{ hoặc } =SUM(\text{Vùng})")
        st.write("**Công dụng:** Trả về tổng của danh sách các đối số hoặc vùng chứa giá trị số.")
        st.code("Ví dụ: =SUM(3, 1+2, A1, A1:C1) -> Kết quả là tổng các giá trị trong danh sách.")

        # Hàm AVERAGE
        st.markdown("#### 2. Hàm AVERAGE")
        st.latex(r"=AVERAGE(X_1, X_2, ..., X_{255}) \text{ hoặc } =AVERAGE(\text{Vùng})")
        st.write("**Công dụng:** Trả về giá trị trung bình cộng của các đối số hoặc ô chứa số trong vùng.")
        st.code("Ví dụ: =AVERAGE(3, 2+1, A1:C2) -> Tính trung bình cộng các số và vùng được chọn.")

        # Hàm COUNT/COUNTA
        st.markdown("#### 3. Hàm COUNT & COUNTA")
        st.write("**COUNT**: Đếm các ô chứa giá trị kiểu số (không đếm ô là chuỗi và ô rỗng).")
        st.write("**COUNTA**: Đếm các ô chứa dữ liệu bất kỳ (số và chuỗi), không đếm ô rỗng.")
        st.latex(r"=COUNT(\text{Vùng}) \quad | \quad =COUNTA(\text{Vùng})")

        # Hàm SUMIFS
        st.markdown("#### 4. Hàm SUMIFS")
        st.latex(r"=SUMIFS(\text{Vùng tính tổng}, \text{Vùng ĐK 1}, \text{ĐK 1}, \text{Vùng ĐK 2}, \text{ĐK 2}, ...)")
        st.write("**Công dụng:** Tính tổng các ô thỏa mãn nhiều điều kiện đồng thời.")
        st.code("Ví dụ: =SUMIFS(F5:F12, C5:C12, 'LA', D5:D12, 'Cam 1') -> Tính tổng tiền mặt hàng 'Cam 1' tại cửa hàng 'LA'.")

    # --- NHÓM HÀM CHUỖI ---
    with st.container():
        st.info("🔤 **NHÓM HÀM XỬ LÝ CHUỖI**")
        
        st.markdown("#### 1. Hàm LEFT & RIGHT")
        st.latex(r"=LEFT(\text{Chuỗi}, n) \quad | \quad =RIGHT(\text{Chuỗi}, n)")
        st.write("**LEFT**: Trích n ký tự của chuỗi tính từ trái qua.")
        st.write("**RIGHT**: Trích n ký tự của chuỗi tính từ phải qua.")
        st.code("Ví dụ: =LEFT('Cao Đẳng', 3) -> Kết quả: 'Cao'.")

        st.markdown("#### 2. Hàm MID")
        st.latex(r"=MID(\text{Chuỗi}, m, n)")
        st.write("**Công dụng:** Trả về chuỗi con gồm n ký tự, bắt đầu từ vị trí m tính từ trái sang.")
        st.code("Ví dụ: =MID('Cao Đẳng Thương Mại', 10, 6) -> Kết quả: 'Thương'.")

        st.markdown("#### 3. Hàm VALUE")
        st.latex(r"=VALUE(\text{Chuỗi số})")
        st.write("**Công dụng:** Chuyển đổi chuỗi ký tự chứa số thành giá trị số để tính toán.")

    # --- NHÓM HÀM LOGIC ---
    with st.container():
        st.info("⚖️ **NHÓM HÀM LOGIC & ĐIỀU KIỆN**")
        
        st.markdown("#### 1. Hàm IF")
        st.latex(r"=IF(\text{Điều kiện}, \text{Giá trị 1}, \text{Giá trị 2})")
        st.write("**Công dụng:** Kiểm tra điều kiện. Nếu Đúng trả về Giá trị 1, nếu Sai trả về Giá trị 2.")
        st.code("Ví dụ: =IF(B2>=5, 'Đậu', 'Trượt') -> Nếu điểm >= 5 thì hiện 'Đậu'.")

        st.markdown("#### 2. Hàm AND & OR")
        st.write("**AND**: Trả về TRUE nếu **tất cả** các điều kiện đều đúng.")
        st.write("**OR**: Trả về TRUE nếu có **ít nhất một** điều kiện đúng.")

    # --- NHÓM HÀM DÒ TÌM ---
    with st.container():
        st.info("🔍 **NHÓM HÀM DÒ TÌM (LOOKUP FUNCTIONS)**")
        
        # 1. Hàm VLOOKUP
        st.markdown("#### 1. Hàm VLOOKUP")
        st.latex(r"=VLOOKUP(\text{Giá trị dò tìm}, \text{Vùng dò tìm}, m, n)")
        st.write("**Công dụng:** Hàm thực hiện lấy giá trị dò tìm so sánh từ trên xuống dưới các giá trị trong cột đầu tiên của vùng dò tìm.")
        st.markdown("""
        * **Giá trị dò tìm**: Giá trị dùng để dò tìm, được so sánh với các giá trị ở cột đầu tiên của bảng dò tìm.
        * **Vùng dò tìm**: Phạm vi dò tìm, phải chứa cột đầu tiên (chứa giá trị dò) và cột chứa giá trị trả về.
        * **m (cột tham chiếu)**: Số thứ tự của cột cần lấy dữ liệu trả về trên vùng dò tìm.
        * **n (cách dò tìm)**: Nhận giá trị 0 (dò tìm chính xác) hoặc 1 (dò tìm tương đối).
        """)
        st.code("Ví dụ: Dò tìm chức vụ của nhân viên dựa trên mã nhân viên trong bảng dọc.")
        # Chèn ảnh minh họa ví dụ từ tài liệu
        try:
            st.image("images/vlookup_VD.png", caption="Hình 3.5: Ví dụ minh họa hàm VLOOKUP tra cứu Chức vụ")
        except:
            st.warning("⚠️ Chú ý: Hãy đảm bảo file ảnh đã được tải lên.")

        st.write("**Công dụng:** Hàm thực hiện lấy giá trị dò tìm so sánh từ trên xuống dưới các giá trị trong cột đầu tiên của vùng dò tìm.")
        
        # Phân tích chi tiết tham số cho học sinh
        st.markdown("""
        **Phân tích các thành phần trong ví dụ trên:**
        * **Giá trị dò tìm**: Là ô `A3` (Mã nhân viên 'HT').
        * **Vùng dò tìm**: Là phạm vi `E2:F6` (Bảng chức vụ bên phải, lưu ý cố định bảng dò để copy công thức.).
        * **m (cột tham chiếu)**: Số `2` (Vì tên Chức vụ nằm ở cột thứ 2 trong bảng dò tìm).
        * **n (cách dò tìm)**: Số `0` (Để tìm chính xác tuyệt đối mã nhân viên).
        * **Vậy hàm viết trong ô C3 sẽ là: =VLOOKUP(A3,\$E\$2:\$F\$6,2,0)**
        """)

        # 2. Hàm HLOOKUP
        st.markdown("#### 2. Hàm HLOOKUP")
        st.latex(r"=HLOOKUP(\text{Giá trị dò tìm}, \text{Vùng dò tìm}, m, n)")
        st.write("**Công dụng:** Hàm thực hiện lấy giá trị dò tìm so sánh từ trái qua phải các giá trị trong hàng đầu tiên của vùng dò tìm.")
        st.markdown("""
        * **Giá trị dò tìm**: Giá trị dùng để dò tìm, được so sánh với các giá trị ở hàng đầu tiên của bảng dò tìm.
        * **Vùng dò tìm**: Phạm vi dò tìm, phải chứa hàng đầu tiên (chứa giá trị dò) và hàng chứa giá trị trả về.
        * **m (hàng tham chiếu)**: Số thứ tự của hàng cần lấy dữ liệu trả về trên vùng dò tìm.
        * **n (cách dò tìm)**: Nhận giá trị 0 (dò tìm chính xác) hoặc 1 (dò tìm tương đối).
        """)
        st.code("Ví dụ: Dò tìm chức vụ của nhân viên dựa trên mã nhân viên trong bảng ngang.")
        # Chèn ảnh minh họa ví dụ từ tài liệu
        try:
            st.image("images/hlookup_VD.png", caption="Hình: Ví dụ minh họa hàm HLOOKUP tra cứu Chức vụ")
        except:
            st.warning("⚠️ Chú ý: Hãy đảm bảo file ảnh đã được tải lên.")

        st.write("**Công dụng:** Hàm thực hiện lấy giá trị dò tìm so sánh từ trên xuống dưới các giá trị trong cột đầu tiên của vùng dò tìm.")
        
        # Phân tích chi tiết tham số cho học sinh
        st.markdown("""
        **Phân tích các thành phần trong ví dụ trên:**
        * **Giá trị dò tìm**: Là ô `A3` (Mã nhân viên 'HT').
        * **Vùng dò tìm**: Là phạm vi `A10:E11` (Bảng chức vụ bên dưới, lưu ý cố định bảng dò để copy công thức.).
        * **m (cột tham chiếu)**: Số `2` (Vì tên Chức vụ nằm ở cột thứ 2 trong bảng dò tìm).
        * **n (cách dò tìm)**: Số `0` (Để tìm chính xác tuyệt đối mã nhân viên).")
        * **Vậy hàm viết trong ô C3 sẽ là: =HLOOKUP(A3,\$A\$10:\$E\$11,2,0)**
        """)
        # Lưu ý về lỗi tham chiếu
        st.warning("""
        **Lưu ý về một số thông báo lỗi:**
        * **#VALUE!**: Xuất hiện nếu tham số m < 1.
        * **#REF!**: Xuất hiện nếu tham số m lớn hơn số cột/hàng trong bảng dò tìm.
        * **#N/A**: Xuất hiện khi không tìm thấy giá trị trùng khớp (với cách dò tìm n = 0).
        """)

    st.divider()
    st.success("🏁 Bạn đã xem xong phần lý thuyết chi tiết về các hàm cơ bản trong Excel.")
# --- HÀM CHO PHẦN THỰC HÀNH ---


def hien_thi_thuc_hanh():
    st.header("🛠 PHẦN THỰC HÀNH: BÀI TẬP HÀM THỐNG KÊ")
    st.write("---")

    # 1. Hiển thị đề bài (Lớp Dữ liệu & Giao diện)
    #st.subheader("Bài 1")
    #try:
        # Đảm bảo bạn đã upload file 'HamTK.png' lên GitHub cùng file code
        #st.image("HamTK.png", caption="Yêu cầu: Điền công thức cho các ô trống (từ câu 1 đến câu 5)")
    #except FileNotFoundError:
        #st.error("⚠️ Không tìm thấy file 'HamTK.png'. Hãy đảm bảo bạn đã tải ảnh lên GitHub cùng thư mục với file code.")
        #st.info("Trong lúc chờ tải ảnh, bạn có thể xem bảng dữ liệu mô phỏng bên dưới để làm bài.")

    # 2. Mô phỏng bảng dữ liệu bằng Pandas để chấm điểm chính xác
    # Việc này giúp hệ thống hiểu được ngữ cảnh địa chỉ ô (A1, B2...)
    data = {
        'A': ['Mã số', 'A01', 'A02', 'B01', 'B02', 'A03'],
        'B': ['Tên hàng', 'Bột giặt LIX', 'Dầu ăn Tường An', 'Gạo nàng hương', 'Sữa đặc Cô gái HL', 'Bột giặt OMO'],
        'C': ['Số lượng', 100, 150, 200, 250, 300],
        'D': ['Đơn giá', 15000, 25000, 18000, 20000, 45000],
        'E': ['Thành tiền', 1500000, 3750000, 3600000, 5000000, 13500000]
    }
    st.subheader("Đề bài")
    # Tạo DataFrame với chỉ số dòng bắt đầu từ 1 để giống Excel
    df = pd.DataFrame(data)
    df.index = df.index + 1 
    
    with st.expander("👀 Xem bảng dữ liệu dạng số (để đối chiếu ô)"):
        st.dataframe(df)

    st.write("---")
    st.subheader("Trả lời. Nhập công thức của bạn")
    st.markdown("*(Lưu ý: Công thức phải bắt đầu bằng dấu `=` và không chứa khoảng trắng)*")

    # Khởi tạo trạng thái để theo dõi số câu đúng (dùng cho nút Hoàn thành)
    if 'correct_answers' not in st.session_state:
        st.session_state.correct_answers = 0

    # Hàm tiện ích để chuẩn hóa và kiểm tra công thức
    def check_formula(user_input, correct_formulas, input_key):
        if not user_input:
            return # Chưa nhập thì không làm gì
        
        # Chuẩn hóa: Viết hoa, xóa khoảng trắng
        normalized_input = user_input.strip().upper().replace(" ", "")
        
        if normalized_input in correct_formulas:
            st.success("✅ Tuyệt vời! Bạn đã nắm vững kiến thức.")
            # Đánh dấu câu này đã trả lời đúng (nếu chưa được đánh dấu)
            if f"{input_key}_correct" not in st.session_state:
                st.session_state[f"{input_key}_correct"] = True
                st.session_state.correct_answers += 1
        else:
            st.error("❌ Công thức chưa chính xác.")
            # Phản hồi sư phạm dựa trên lỗi thường gặp
            if not normalized_input.startswith("="):
                st.warning("💡 Gợi ý: Công thức trong Excel luôn phải bắt đầu bằng dấu `=`")
            elif "SUM" not in normalized_input and input_key in ['q1', 'q2', 'q3']:
                 st.warning("💡 Gợi ý: Yêu cầu tính 'Tổng', hãy xem lại hàm SUM.")
            elif "COUNTA" not in normalized_input and input_key == 'q4':
                 st.warning("💡 Gợi ý: Yêu cầu đếm 'Mặt hàng' (dữ liệu dạng chuỗi), hãy xem lại hàm COUNTA.")
            elif "MAX" not in normalized_input and input_key == 'q5':
                 st.warning("💡 Gợi ý: Yêu cầu tìm giá trị 'Cao nhất', hãy xem lại hàm MAX.")
            else:
                 st.warning(f"💡 Gợi ý: Hãy kiểm tra kỹ vùng dữ liệu tham chiếu (ví dụ: A2:A6, C2:C6...).")

    # --- Khu vực nhập liệu cho từng câu ---
    
    # Câu 1
    st.markdown("**1. Tính tổng số lượng hàng:**")
    f_c7 = st.text_input("Nhập công thức:", key="q1_input")
    check_formula(f_c7, ["=SUM(C2:C6)", "=SUM(C2,C3,C4,C5,C6)"], "q1")

    # Câu 2
    st.markdown("**2. Tính tổng thành tiền mặt hàng A:**")
    f_e8 = st.text_input("Nhập công thức ô:", key="q2_input")
    # Sử dụng Regex để chấp nhận cả SUMIF và SUMIFS
    check_formula(f_e8, ["=SUMIF(A2:A6,\"A*\",E2:E6)", "=SUMIFS(E2:E6,A2:A6,\"A*\")"], "q2")

    # Câu 3
    st.markdown("**3. Tính tổng thành tiền mặt hàng B:**")
    f_e9 = st.text_input("Nhập công thức ô:", key="q3_input")
    check_formula(f_e9, ["=SUMIF(A2:A6,\"B*\",E2:E6)", "=SUMIFS(E2:E6,A2:A6,\"B*\")"], "q3")

    # Câu 4
    st.markdown("**4. Có bao nhiêu mặt hàng:**")
    f_b10 = st.text_input("Nhập công thức ô:", key="q4_input")
    check_formula(f_b10, ["=COUNTA(B2:B6)", "=COUNT(A2:A6)"], "q4") # Chấp nhận COUNT nếu đếm mã số

    # Câu 5
    st.markdown("**5. Đơn giá cao nhất là:**")
    f_d11 = st.text_input("Nhập công thức:", key="q5_input")
    check_formula(f_d11, ["=MAX(D2:D6)"], "q5")

    st.write("---")
    
    # Nút Hoàn thành và bóng bay
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if st.button("📊 Hoàn thành bài thực hành", use_container_width=True):
            if st.session_state.correct_answers == 5:
                st.success(f"🎉 Chúc mừng! Bạn đã hoàn thành xuất sắc 5/5 câu hỏi.")
                st.balloons() # Hiệu ứng bóng bay
            else:
                st.warning(f"Bạn mới hoàn thành {st.session_state.correct_answers}/5 câu. Hãy cố gắng trả lời đúng hết các câu nhé!")
def hien_thi_thuc_hanh_ham_if():
    st.header("🛠 PHẦN THỰC HÀNH: HÀM IF & LOGIC")
    st.write("---")

    # 1. Hiển thị đề bài
    #st.subheader("I. Đề bài: Kết quả kiểm tra chất lượng")
    #try:
      #  st.image("HamIF.png", caption="Yêu cầu: Điền công thức cho cột Kết quả, Xếp loại và Học bổng")
    #except:
        #st.warning("⚠️ Hãy đảm bảo bạn đã tải file ảnh 'HamIF.png' lên GitHub.")

    # 2. Bảng dữ liệu mô phỏng để sinh viên đối chiếu địa chỉ ô [cite: 51, 185]
    st.markdown("**Bảng dữ liệu mô phỏng (Địa chỉ ô tương ứng):**")
    data_if = {
        'A': ['Họ tên', 'Anh', 'Thảo', 'Việt', 'Lan', 'Thành'],
        'B': ['KT', 6, 8, 9, 5, 9],
        'C': ['CT', 4, 5, 8, 10, 10],
        'D': ['Tin', 4, 10, 5, 6, 8.5],
        'E': ['ĐTB', '', '', '', '', ''],
        'F': ['Kết quả', '', '', '', '', ''],
        'G': ['Xếp loại', '', '', '', '', ''],
        'H': ['Học bổng', '', '', '', '', '']
    }
    df_if = pd.DataFrame(data_if)
    df_if.index = df_if.index + 3 # Bắt đầu từ dòng 2 theo đề bài [cite: 190]
    st.dataframe(df_if)

    st.write("---")
    st.subheader("Trả lời: Nhập công thức của bạn (Tính cho dòng đầu tiên - dòng 4)")

    # Khởi tạo trạng thái hoàn thành [cite: 69]
    if 'correct_if' not in st.session_state:
        st.session_state.correct_if = 0

    def check_if_formula(user_input, correct_options, key_name, hint_text):
        if user_input:
            # Chuẩn hóa chuỗi: viết hoa, xóa khoảng trắng [cite: 63]
            processed_input = user_input.strip().upper().replace(" ", "")
            if processed_input in [opt.upper().replace(" ", "") for opt in correct_options]:
                st.success("✅ Tuyệt vời! Bạn đã nắm vững kiến thức.")
                if f"{key_name}_done" not in st.session_state:
                    st.session_state[f"{key_name}_done"] = True
                    st.session_state.correct_if += 1
            else:
                st.error(f"❌ Chưa chính xác. Gợi ý: {hint_text}[cite: 73].")

    # --- Câu 1: Tính ĐTB ---
    st.markdown("**Câu 1: Tính ĐTB 3 môn (Ô E4):**")
    f_e4 = st.text_input("Nhập công thức ô E4:", key="if_q1")
    check_if_formula(f_e4, ["=AVERAGE(B4:D4)", "=(B4+C4+D4)/3"], "q1", 
                     "Sử dụng hàm AVERAGE hoặc tính tổng 3 môn rồi chia 3.")

    # --- Câu 2: Kết quả ---
    st.markdown("**Câu 2: Nếu ĐTB > 5 là 'Đạt', ngược lại 'Thi lại' (Ô F4):**")
    f_f4 = st.text_input("Nhập công thức ô F4:", key="if_q2")
    check_if_formula(f_f4, ["=IF(E4>5,\"Đạt\",\"Thi lại\")"], "q2", 
                     "Sử dụng hàm IF với điều kiện E4 > 5.")

    # --- Câu 3: Xếp loại ---
    st.markdown("**Câu 3: Xếp loại dựa trên ĐTB:Nếu ĐTB>=8.5 xếp loại A, ĐTB>=7 xếp loại B Còn lại loại C(Ô G4):**")
    f_g4 = st.text_input("Nhập công thức ô G4:", key="if_q3")
    check_if_formula(f_g4, ["=IF(E4>=8.5,\"A\",IF(E4>=7,\"B\",\"C\"))"], "q3", 
                     "Sử dụng hàm IF lồng nhau để xét nhiều điều kiện.")

    # --- Câu 4: Học bổng ---
    st.markdown("**Câu 4: Học bổng nếu Kết quả 'Đạt' VÀ ĐTB >= 9 được nhận : 3.000.000 còn lại để trống ô(Ô H4):**")
    f_h4 = st.text_input("Nhập công thức ô H4:", key="if_q4")
    check_if_formula(f_h4, ["=IF(AND(F4=\"Đạt\",E4>=9),3000000,\"\")", "=IF(AND(E4>=9,F4=\"Đạt\"),3000000,\"\")"], "q4", 
                     "Sử dụng hàm IF kết hợp hàm AND. Lưu ý để trống ô thì dùng cặp dấu nháy kép rỗng \"\".")

    # Nút Hoàn thành [cite: 74]
    if st.button("📊 Hoàn thành phần thực hành Hàm IF"):
        if st.session_state.correct_if >= 4:
            st.success("🎉 Chúc mừng! Bạn đã hoàn thành xuất sắc các bài tập về Hàm IF!")
            st.balloons()
        else:
            st.warning(f"Bạn đã làm đúng {st.session_state.correct_if}/4 câu. Hãy hoàn thiện các câu còn lại!")

def hien_thi_thuc_hanh_lookup():
    st.header("🛠 PHẦN THỰC HÀNH: HÀM DÒ TÌM VÀ XỬ LÝ CHUỖI")
    st.write("---")

    # 1. Hiển thị đề bài
    st.subheader("Đề bài:")
    try:
        st.image("images/HamLookup.png", caption="Yêu cầu: Sử dụng VLOOKUP và HLOOKUP kết hợp hàm LEFT, RIGHT") [cite: 1471, 1472, 1568, 1598]
    except:
        st.warning("⚠️ Hãy đảm bảo bạn đã tải file ảnh 'HamLookup.png' lên GitHub.")

    # 2. Mô phỏng bảng dữ liệu để học sinh đối chiếu địa chỉ ô
    #st.markdown("**Bảng dữ liệu mô phỏng (Địa chỉ ô tương ứng):**")
    # Bảng 1 (Bảng chính cần điền dữ liệu) [cite: 26, 27]
    #data_lookup = {
        #'A': ['Mã hàng', 'BLXK', 'DLXK', 'MLTN', 'BLTN'],
        #'B': ['Tên hàng', '', '', '', ''],
        #'C': ['Giá', '', '', '', ''],
        #'D': ['Phân phối', '', '', '', ''],
        #'E': ['% Thuế', '', '', '', '']
    #}
    #df_main = pd.DataFrame(data_lookup)
    #df_main.index = df_main.index + 2 # Dòng 2 đến dòng 6 [cite: 354, 355]
    st.write("Bảng 1 (Vùng A2:E6):")
    #st.dataframe(df_main)

    st.write("---")
    st.subheader("Trả lời: Nhập công thức cho dòng đầu tiên (Dòng 3)")

    # Khởi tạo trạng thái hoàn thành bài học [cite: 74]
    if 'score_lookup' not in st.session_state:
        st.session_state.score_lookup = 0

    def check_lookup_formula(user_input, correct_options, key_name, hint_text):
        if user_input:
            processed = user_input.strip().upper().replace(" ", "")
            if processed in [opt.upper().replace(" ", "") for opt in correct_options]:
                st.success("✅ Tuyệt vời! Bạn đã nắm vững kiến thức.") 
                if f"{key_name}_done" not in st.session_state:
                    st.session_state[f"{key_name}_done"] = True
                    st.session_state.score_lookup += 1
            else:
                st.error(f"❌ Chưa chính xác. Gợi ý: {hint_text}.") 

    # --- Yêu cầu 1: Tên hàng và Giá (Sử dụng VLOOKUP + LEFT) --- [cite: 1472, 1568]
    # Yêu cầu 1: Tên hàng (VLOOKUP + LEFT) [cite: 1472, 1570]
    st.markdown("**1. Tên hàng (Ô B3) - Dựa vào kí tự đầu và Bảng 2:**")
    f_b3 = st.text_input("Nhập công thức ô B3:", key="lookup_q1")
    
    if f_b3:
        processed_b3 = f_b3.strip().upper().replace(" ", "")
        if processed_b3 == "=VLOOKUP(LEFT(A3,1),$A$9:$C$11,2,0)" or processed_b3 == "=VLOOKUP(LEFT(A3,1),$A$8:$C$11,2,0)":
            st.success("✅ Tuyệt vời! Bạn đã nắm vững kiến thức.")
        else:
            st.error("❌ Sai rồi! Hãy thử nhập lại.")
            st.info("💡 Gợi ý: Dùng LEFT(A3,1) để lấy ký tự đầu làm giá trị dò tìm.Lưu ý cố định vùng dò")
            
    st.markdown("**2. Giá (Ô C3) - Dựa vào kí tự đầu và Bảng 2:**")
    f_c3 = st.text_input("Nhập công thức ô C3:", key="lookup_q2")
    check_lookup_formula(f_c3, ["=VLOOKUP(LEFT(A3,1),$A$9:$C$11,3,0)","=VLOOKUP(LEFT(A3,1),$A$8:$C$11,3,0)"], "q2", 
                         "Tương tự Tên hàng nhưng lấy dữ liệu ở cột 3 của Bảng 2.Lưu ý cố định vùng dò") 

    # --- Yêu cầu 2: Phân phối và Thuế (Sử dụng HLOOKUP + RIGHT) --- [cite: 1477, 1598]
    st.markdown("**3. Phân phối (Ô D3) - Dựa vào 2 kí tự cuối và Bảng 3:**")
    f_d3 = st.text_input("Nhập công thức ô D3:", key="lookup_q3")
    check_lookup_formula(f_d3, ["=HLOOKUP(RIGHT(A3,2),$E$8:$G$10,2,0)","=HLOOKUP(RIGHT(A3,2),$F$8:$G$10,2,0)"], "q3", 
                         "Sử dụng HLOOKUP kết hợp RIGHT lấy 2 kí tự cuối mã hàng, dò trong Bảng 3 (E8:G10), lấy hàng 2.Lưu ý cố định vùng dò")
    st.markdown("**4. % Thuế (Ô E3) - Dựa vào 2 kí tự cuối và Bảng 3:**")
    f_e3 = st.text_input("Nhập công thức ô E3:", key="lookup_q4")
    check_lookup_formula(f_e3, ["=HLOOKUP(RIGHT(A3,2),$E$8:$G$10,3,0)","=HLOOKUP(RIGHT(A3,2),$F$8:$G$10,3,0)"], "q4", 
                         "Tương tự Phân phối nhưng lấy dữ liệu ở hàng 3 của Bảng 3.Lưu ý cố định vùng dò")

    # Nút Hoàn thành [cite: 74]
    if st.button("📊 Hoàn thành phần thực hành Dò tìm"):
        if st.session_state.score_lookup >= 4:
            st.success("🎉 Chúc mừng! Bạn đã hoàn thành xuất sắc bài tập VLOOKUP và HLOOKUP!")
            st.balloons() 
        else:
            st.warning(f"Bạn đã làm đúng {st.session_state.score_lookup}/4 câu. Hãy kiểm tra lại các phần gợi ý nếu có câu sai nhé!")
    # Để chạy, bạn chỉ cần gọi hàm này trong file chính:
    
# --- HÀM CHO PHẦN TRẮC NGHIỆM ---
def hien_thi_trac_nghiem():
    st.header("✍️ KIỂM TRA KIẾN THỨC EXCEL")
    st.write("Hãy chọn đáp án đúng nhất cho các câu hỏi dưới đây:")

    # Khởi tạo danh sách câu hỏi từ file soạn sẵn
    questions = [
        {
            "id": 1,
            "question": "Tại ô A2 có giá trị 25; Tại ô B2 gõ công thức $=SQRT(A2)$ thì kết quả là:",
            "options": ["0", "5", "#VALUE!", "#NAME!"],
            "answer": "5",
            "explain": "Hàm SQRT dùng để tính căn bậc hai của một số. Căn bậc hai của 25 là 5."
        },
        {
            "id": 2,
            "question": "Kết quả của công thức: $=IF(3>5, 100, IF(5<6, 200, 300))$ là:",
            "options": ["200", "100", "300", "False"],
            "answer": "200",
            "explain": "3>5 là sai, hàm xét tiếp hàm IF thứ hai. 5<6 là đúng nên trả về giá trị 200."
        },
        {
            "id": 3,
            "question": "Ô D2 có công thức $=B2*C2/100$. Nếu sao chép đến ô G6 thì công thức sẽ là:",
            "options": ["E7*F7/100", "B6*C6/100", "E6*F6/100", "E2*C2/100"],
            "answer": "E6*F6/100",
            "explain": "Đây là địa chỉ tương đối. Khi dịch chuyển từ cột D sang G (3 cột) và dòng 2 sang 6 (4 dòng), các tham chiếu sẽ dịch chuyển tương ứng."
        },
        {
            "id": 4,
            "question": "Để sửa dữ liệu trong một ô tính mà không cần nhập lại, ta thực hiện:",
            "options": ["Bấm phím F2", "Bấm phím F4", "Bấm phím F10", "Bấm phím F12"],
            "answer": "Bấm phím F2",
            "explain": "Phím F2 cho phép chuyển ô tính sang chế độ hiệu chỉnh (Edit mode)."
        },
        {
            "id": 5,
            "question": "Trong Excel, khi viết sai tên hàm trong tính toán, chương trình thông báo lỗi?",
            "options": ["#NAME!", "#VALUE!", "#N/A!", "#DIV/0!"],
            "answer": "#NAME!",
            "explain": "Lỗi #NAME? xuất hiện khi Excel không nhận diện được văn bản trong công thức (thường do viết sai tên hàm)."
        },
        {
            "id": 6,
            "question": "Muốn sắp xếp danh sách dữ liệu theo thứ tự tăng (giảm), ta thực hiện:",
            "options": ["Review - Sort", "View - Sort", "Data - Sort", "Page Layout - Sort"],
            "answer": "Data - Sort",
            "explain": "Công cụ Sort nằm trong thẻ Data trên thanh Ribbon."
        },
        {
            "id": 7,
            "question": "Tại ô D5 gõ công thức $=MOD(22, 7)$ sẽ cho kết quả là:",
            "options": ["2", "3", "1", "0"],
            "answer": "1",
            "explain": "Hàm MOD trả về số dư của phép chia. 22 chia 7 được 3 dư 1."
        },
        {
            "id": 8,
            "question": "Kết quả của biểu thức $=PROPER(\" hoang lien son\") là:",
            "options": ["HOANG LIEN SON", "hoang lien son", "Hoang Lien Son", "HoangLienSon"],
            "answer": "Hoang Lien Son",
            "explain": "Hàm PROPER viết hoa chữ cái đầu tiên của mỗi từ trong chuỗi văn bản."
        },
        {
            "id": 9,
            "question": "Kết quả của công thức: $=4/2^3$ là:",
            "options": ["8", "0.5", "6", "5"],
            "answer": "0.5",
            "explain": "Phép tính lũy thừa thực hiện trước: $2^3 = 8$. Sau đó $4/8 = 0.5$."
        },
        {
            "id": 10,
            "question": "Tại ô F12 có công thức: $=\"Cầu Rồng. \" \& MIN(2013, 2015)$ thì kết quả là:",
            "options": ["FALSE", "Cầu Rồng. Min(2013.2015)", "Cầu Rồng.2015", "Cầu Rồng. 2013"],
            "answer": "Cầu Rồng. 2013",
            "explain": "Toán tử & dùng để nối chuỗi. Hàm MIN(2013, 2015) trả về 2013."
        },
        {
            "id": 11,
            "question": "Phím tắt để tính tổng nhanh (Auto Sum) các ô liên tục của một cột là:",
            "options": ["Ctrl + =", "Data / Subtotals", "Alt + =", "Tất cả đều đúng"],
            "answer": "Alt + =",
            "explain": "Tổ hợp phím Alt + = giúp tự động tạo hàm SUM cho các ô dữ liệu lân cận."
        },
        {
            "id": 12,
            "question": "Trong các biểu thức sau, biểu thức nào có kết quả là FALSE?",
            "options": ["OR(5>4, 10>20)", "AND(50>6, OR(10>6, 1>3))", "OR(AND(5<4, 3>1), 10>20)", "AND(5>4, 3>1, 30>20)"],
            "answer": "OR(AND(5<4, 3>1), 10>20)",
            "explain": "AND(5<4, 3>1) là FALSE. 10>20 là FALSE. OR(FALSE, FALSE) trả về FALSE."
        }
    ]

    # Tạo form để lưu kết quả làm bài
    with st.form("quiz_form"):
        user_answers = {}
        for q in questions:
            st.write(f"**Câu {q['id']}:** {q['question']}")
            user_answers[q['id']] = st.radio(f"Chọn đáp án cho câu {q['id']}:", 
                                             q['options'], 
                                             key=f"q_{q['id']}", 
                                             label_visibility="collapsed")
            st.write("")

        submit_button = st.form_submit_button("Hoàn thành")

    if submit_button:
        score = 0
        wrong_answers = []

        for q in questions:
            if user_answers[q['id']] == q['answer']:
                score += 1
            else:
                wrong_answers.append(q)

        # Hiển thị kết quả
        st.subheader(f"📊 Kết quả: {score}/{len(questions)} câu đúng")
        
        if score == len(questions):
            st.success("Tuyệt vời! Bạn đã nắm vững kiến thức")
            st.balloons() # Hiệu ứng bóng bay khi kết thúc [cite: 74]
        else:
            st.warning("Bạn cần xem lại các câu sau:")
            for w in wrong_answers:
                with st.expander(f"Giải thích câu {w['id']}"):
                    st.write(f"**Câu hỏi:** {w['question']}")
                    st.write(f"**Đáp án đúng:** {w['answer']}")
                    st.write(f"**Giải thích:** {w['explain']}")

    # Nút thử lại
    if st.button("Thử lại"):
        st.rerun() # Làm lại từ đầu

# --- GIAO DIỆN CHÍNH (MAIN) ---
def main():
    # 1. Thanh bên (Sidebar) để chọn nội dung 
    st.sidebar.title("Mục lục học tập")
    lua_chon = st.sidebar.radio(
        "Chọn phần bạn muốn học:",
        ["Phần Lý thuyết", "Phần thực hành", "Phần trắc nghiệm"]
    )

    # 2. Tiêu đề và mô tả trang chính [cite: 59]
    st.title("🚀 Ứng dụng học Microsoft Excel tương tác")
    st.write("Chào mừng bạn đến với hệ thống học tập thông minh. Ứng dụng này giúp bạn hệ thống hóa kiến thức và luyện tập các kỹ năng Excel cơ bản.")
    
    # 3. Điều hướng nội dung dựa trên lựa chọn ở Sidebar
    if lua_chon == "Phần Lý thuyết":
        hien_thi_ly_thuyet()
    elif lua_chon == "Phần thực hành":
        hien_thi_thuc_hanh()
        hien_thi_thuc_hanh_ham_if()
        hien_thi_thuc_hanh_lookup()
    elif lua_chon == "Phần trắc nghiệm":
        hien_thi_trac_nghiem()

if __name__ == "__main__":
    main()
