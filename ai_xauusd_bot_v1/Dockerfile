FROM python:3.10-slim

# Cài các gói hệ thống cần thiết cho tkinter
RUN apt-get update && apt-get install -y \
    python3-tk \
    tk \
    tcl \
    libx11-6 \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Tạo thư mục làm việc
WORKDIR /app

# Sao chép mã nguồn vào container
COPY . .

# Cài đặt các thư viện Python
RUN pip install --no-cache-dir -r requirements.txt

# Mặc định chạy chương trình
CMD ["python", "translate_code_design.py"]
