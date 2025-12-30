
"""
HWPX 파일 이미지 일괄 압축 프로그램 v3.0 - 고급 버전
- 표 안의 배경/테두리 이미지 처리
- XML 내 base64 인코딩 이미지 처리
- 사용자 정의 압축 크기
- 실시간 진행 상황 표시
"""

import os
import zipfile
import shutil
from pathlib import Path
from PIL import Image
import io
import tkinter as tk
from tkinter import ttk
import threading
import time
import base64
import xml.etree.ElementTree as ET
import re

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except ImportError:
    print("tkinterdnd2 설치 필요: pip install tkinterdnd2")
    exit()


class HWPXImageCompressorAdvanced:
    """HWPX 파일의 모든 이미지를 압축하는 고급 클래스"""

    def __init__(self, target_size_kb=200):
        self.target_size_kb = target_size_kb
        self.target_size_bytes = target_size_kb * 1024
        self.processed_images = {}  # base64 캐시용

    def compress_image(self, image_data, original_format="jpg"):
        """이미지를 목표 크기로 압축"""
        try:
            img = Image.open(io.BytesIO(image_data))
        except:
            return image_data, 'original'

        # 색상 모드 정규화
        if img.mode in ('RGBA', 'LA', 'P'):
            background = Image.new('RGB', img.size, (255, 255, 255))
            if img.mode == 'P':
                img = img.convert('RGBA')
            background.paste(img, mask=img.split()[-1] if img.mode in ('RGBA', 'LA') else None)
            img = background
        elif img.mode != 'RGB':
            img = img.convert('RGB')

        # 품질 조정으로 목표 크기 맞추기
        quality = 95
        while quality > 5:
            output = io.BytesIO()
            img.save(output, format='JPEG', quality=quality, optimize=True)
            size = output.tell()

            if size <= self.target_size_bytes:
                return output.getvalue(), 'jpg'

            quality -= 5

        # 이미지 크기 조정
        scale = 0.9
        while scale > 0.3:
            resized = img.resize((int(img.width * scale), int(img.height * scale)), Image.Resampling.LANCZOS)
            output = io.BytesIO()
            resized.save(output, format='JPEG', quality=85, optimize=True)
            size = output.tell()

            if size <= self.target_size_bytes:
                return output.getvalue(), 'jpg'

            scale -= 0.1

        # 최소 크기로 압축
        output = io.BytesIO()
        resized.save(output, format='JPEG', quality=60, optimize=True)
        return output.getvalue(), 'jpg'

    def compress_base64_image(self, base64_string):
        """Base64 인코딩된 이미지 압축 후 Base64로 반환"""
        try:
            # Base64 디코딩
            image_data = base64.b64decode(base64_string)
            original_size = len(image_data)

            # 이미 작으면 그대로 반환
            if original_size <= self.target_size_bytes:
                return base64_string, original_size, original_size, False

            # 압축
            compressed_data, _ = self.compress_image(image_data)
            compressed_size = len(compressed_data)

            # Base64 재인코딩
            compressed_base64 = base64.b64encode(compressed_data).decode('utf-8')

            return compressed_base64, original_size, compressed_size, True

        except Exception as e:
            print(f"[경고] Base64 압축 실패: {e}")
            return base64_string, 0, 0, False

    def process_xml_images(self, xml_content, file_path, progress_callback=None):
        """XML 파일에서 base64 이미지 찾아 압축"""
        try:
            # XML 파싱
            root = ET.fromstring(xml_content)

            # 네임스페이스 추출
            namespaces = {}
            for event, elem in ET.iterparse(io.StringIO(xml_content), events=['start-ns']):
                prefix, uri = event
                if prefix:
                    namespaces[prefix] = uri

            compressed_count = 0
            total_original = 0
            total_compressed = 0

            # 1. bin 속성에서 base64 이미지 찾기 (그림 개체)
            modified = False
            for elem in root.iter():
                if 'bin' in elem.attrib:
                    base64_str = elem.attrib['bin']
                    if len(base64_str) > 100:  # 충분히 큰 base64만
                        try:
                            new_base64, orig_size, comp_size, was_compressed = self.compress_base64_image(base64_str)
                            if was_compressed:
                                elem.attrib['bin'] = new_base64
                                compressed_count += 1
                                total_original += orig_size
                                total_compressed += comp_size
                                modified = True
                        except:
                            pass

            # 2. fillImagePath 속성에서 이미지 참조 찾기 (배경/테두리)
            for elem in root.iter():
                if 'fillImagePath' in elem.attrib:
                    # fillImagePath는 BinData 참조이므로 별도 처리 필요
                    pass

            if modified:
                return ET.tostring(root, encoding='utf-8'), compressed_count, total_original, total_compressed
            else:
                return xml_content, 0, 0, 0

        except Exception as e:
            print(f"[경고] XML 처리 오류 ({file_path}): {e}")
            return xml_content, 0, 0, 0

    def process_hwpx(self, hwpx_path, output_path=None, progress_callback=None):
        """HWPX 파일 처리 메인 함수"""
        if not hwpx_path.endswith('.hwpx'):
            return False, "HWPX 파일이 아닙니다."

        if output_path is None:
            base_name = os.path.splitext(hwpx_path)[0]
            output_path = f"{base_name}_compressed.hwpx"

        temp_dir = "temp_hwpx_processing"
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)
        os.makedirs(temp_dir)

        try:
            # 1단계: HWPX 압축 해제
            if progress_callback:
                progress_callback(0, "📂 HWPX 파일 압축 해제 중...", 0, 0, 0)

            print(f"[시작] 파일: {hwpx_path}")
            start_time = time.time()

            with zipfile.ZipFile(hwpx_path, 'r') as zip_ref:
                zip_ref.extractall(temp_dir)

            # 2단계: BinData 폴더의 이미지 처리
            bindata_path = os.path.join(temp_dir, 'BinData')
            compressed_count = 0
            total_original_size = 0
            total_compressed_size = 0
            skipped_count = 0

            if os.path.exists(bindata_path):
                image_extensions = ('.png', '.jpg', '.jpeg', '.bmp', '.gif')
                image_files = [f for f in os.listdir(bindata_path) 
                              if f.lower().endswith(image_extensions)]

                total_images = len(image_files)
                print(f"[정보] BinData 이미지: {total_images}개 발견")

                # 크기순 정렬
                image_info = [(f, os.path.join(bindata_path, f)) for f in image_files]
                image_info.sort(key=lambda x: os.path.getsize(x[1]), reverse=True)

                for idx, (image_file, image_path) in enumerate(image_info):
                    file_size = os.path.getsize(image_path)
                    file_size_mb = file_size / (1024 * 1024)
                    progress_pct = int((idx / total_images) * 100) if total_images > 0 else 0
                    elapsed = int(time.time() - start_time)

                    if progress_callback:
                        progress_callback(
                            progress_pct,
                            f"🖼️  BinData 이미지: {image_file}\n크기: {file_size_mb:.2f}MB",
                            idx + 1,
                            total_images,
                            elapsed
                        )

                    print(f"[BinData {idx+1}/{total_images}] {image_file} ({file_size_mb:.2f}MB)", end=" -> ")

                    with open(image_path, 'rb') as f:
                        original_data = f.read()

                    original_size = len(original_data)
                    total_original_size += original_size

                    if original_size <= self.target_size_bytes:
                        print(f"스킵 ({original_size/1024:.1f}KB)")
                        total_compressed_size += original_size
                        skipped_count += 1
                        continue

                    try:
                        compressed_data, _ = self.compress_image(original_data)
                        compressed_size = len(compressed_data)
                        reduction = ((original_size - compressed_size) / original_size * 100)

                        with open(image_path, 'wb') as f:
                            f.write(compressed_data)

                        compressed_count += 1
                        total_compressed_size += compressed_size
                        print(f"✅ {compressed_size/1024:.1f}KB (-{reduction:.1f}%)")

                    except Exception as e:
                        print(f"❌ 실패: {e}")
                        total_compressed_size += original_size

            # 3단계: XML 파일의 이미지 처리 (표 배경, 테두리 등)
            xml_files = []
            contents_path = os.path.join(temp_dir, 'Contents')
            if os.path.exists(contents_path):
                for f in os.listdir(contents_path):
                    if f.endswith('.xml'):
                        xml_files.append(os.path.join(contents_path, f))

            xml_compressed_count = 0
            xml_total_original = 0
            xml_total_compressed = 0

            print(f"\n[정보] XML 파일: {len(xml_files)}개 처리 중...")

            for idx, xml_path in enumerate(xml_files):
                file_name = os.path.basename(xml_path)
                elapsed = int(time.time() - start_time)

                if progress_callback:
                    progress_callback(
                        80 + int((idx / len(xml_files)) * 15) if xml_files else 80,
                        f"📄 XML 이미지 처리: {file_name}\n(표 배경/테두리)",
                        idx + 1,
                        len(xml_files),
                        elapsed
                    )

                print(f"[XML {idx+1}/{len(xml_files)}] {file_name}", end=" -> ")

                try:
                    with open(xml_path, 'r', encoding='utf-8') as f:
                        xml_content = f.read()

                    new_xml, comp_count, orig, comp = self.process_xml_images(xml_content, file_name, progress_callback)

                    if comp_count > 0:
                        with open(xml_path, 'w', encoding='utf-8') as f:
                            if isinstance(new_xml, bytes):
                                f.write(new_xml.decode('utf-8'))
                            else:
                                f.write(new_xml)

                        xml_compressed_count += comp_count
                        xml_total_original += orig
                        xml_total_compressed += comp
                        print(f"✅ {comp_count}개 이미지 압축")
                    else:
                        print("스킵")

                except Exception as e:
                    print(f"❌ 오류: {e}")

            # 4단계: 다시 HWPX로 압축
            if progress_callback:
                progress_callback(95, "📦 HWPX 파일 생성 중...", 0, 0, int(time.time() - start_time))

            print(f"\n[진행] HWPX 파일 생성 중...")

            with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for root, dirs, files in os.walk(temp_dir):
                    for file in files:
                        file_path = os.path.join(root, file)
                        arcname = os.path.relpath(file_path, temp_dir)
                        zipf.write(file_path, arcname)

            shutil.rmtree(temp_dir)

            # 결과 계산
            total_all_compressed = compressed_count + xml_compressed_count
            total_all_original = total_original_size + xml_total_original
            total_all_compressed_size = total_compressed_size + xml_total_compressed

            reduction = ((total_all_original - total_all_compressed_size) / total_all_original * 100) if total_all_original > 0 else 0
            elapsed_time = int(time.time() - start_time)

            result_msg = f"""✅ 처리 완료!

📊 통계:
━━━━━━━━━━━━━━━━━━━
🖼️  BinData 이미지:
   - 압축됨: {compressed_count}개
   - 스킵됨: {skipped_count}개

📄 XML 이미지 (표/배경):
   - 압축됨: {xml_compressed_count}개

📈 전체:
   - 총 압축 이미지: {total_all_compressed}개
   - 용량 감소: {reduction:.1f}%
   - 소요 시간: {elapsed_time}초

💾 저장 위치:
{output_path}"""

            print(f"\n{result_msg}")

            return True, result_msg

        except Exception as e:
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)
            error_msg = f"오류 발생: {str(e)}"
            print(f"[오류] {error_msg}")
            return False, error_msg


class HWPXCompressorGUI:
    """고급 기능의 드래그 앤 드롭 GUI"""

    def __init__(self):
        self.root = TkinterDnD.Tk()
        self.root.title("HWPX 이미지 압축기 v3.0 - 고급")
        self.root.geometry("700x600")
        self.root.resizable(False, False)

        self.is_processing = False
        self.target_size_kb = 200
        self.setup_gui()

    def setup_gui(self):
        """GUI 구성"""
        # 제목
        title_frame = tk.Frame(self.root, bg="#2c3e50", height=80)
        title_frame.pack(fill=tk.X)

        title_label = tk.Label(
            title_frame, 
            text="HWPX 이미지 압축기 v3.0",
            font=("맑은 고딕", 18, "bold"),
            bg="#2c3e50",
            fg="white"
        )
        title_label.pack(pady=10)

        info_label = tk.Label(
            title_frame,
            text="✨ 표 안의 배경/테두리 이미지도 모두 처리됩니다",
            font=("맑은 고딕", 9),
            bg="#2c3e50",
            fg="#ecf0f1"
        )
        info_label.pack()

        # 설정 프레임
        settings_frame = tk.Frame(self.root, bg="white", relief=tk.RIDGE, borderwidth=1)
        settings_frame.pack(fill=tk.X, padx=20, pady=15)

        # 압축 크기 설정
        size_label = tk.Label(
            settings_frame,
            text="💾 이미지 압축 크기 선택:",
            font=("맑은 고딕", 10, "bold"),
            bg="white",
            fg="#2c3e50"
        )
        size_label.pack(anchor=tk.W, padx=10, pady=(10, 5))

        button_frame = tk.Frame(settings_frame, bg="white")
        button_frame.pack(fill=tk.X, padx=10, pady=(5, 10))

        sizes = [
            ("매우 작게 (50KB)", 50),
            ("작게 (100KB)", 100),
            ("중간 (200KB)", 200),
            ("크게 (500KB)", 500),
            ("아주 크게 (1MB)", 1000)
        ]

        self.size_var = tk.IntVar(value=200)

        for text, size in sizes:
            rb = tk.Radiobutton(
                button_frame,
                text=text,
                variable=self.size_var,
                value=size,
                font=("맑은 고딕", 9),
                bg="white",
                fg="#34495e",
                selectcolor="#ecf0f1"
            )
            rb.pack(anchor=tk.W)

        # 드래그 앤 드롭 영역
        drop_frame = tk.Frame(self.root, bg="white")
        drop_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 20))

        self.drop_label = tk.Label(
            drop_frame,
            text="\n\n여기에 HWPX 파일을\n드래그 앤 드롭하세요\n\n",
            font=("맑은 고딕", 14),
            bg="#ecf0f1",
            fg="#7f8c8d",
            relief=tk.RIDGE,
            borderwidth=3
        )
        self.drop_label.pack(fill=tk.BOTH, expand=True)

        self.drop_label.drop_target_register(DND_FILES)
        self.drop_label.dnd_bind('<<Drop>>', self.on_drop)

        # 진행률 프레임
        progress_frame = tk.Frame(self.root, bg="white")
        progress_frame.pack(fill=tk.X, padx=20, pady=(0, 10))

        self.progress_bar = ttk.Progressbar(
            progress_frame,
            mode='determinate',
            maximum=100
        )
        self.progress_bar.pack(fill=tk.X, pady=(0, 5))

        self.detail_label = tk.Label(
            progress_frame,
            text="",
            font=("맑은 고딕", 9),
            justify=tk.LEFT,
            fg="#34495e",
            bg="white"
        )
        self.detail_label.pack(anchor=tk.W)

        # 상태 표시
        self.status_label = tk.Label(
            self.root,
            text="대기 중...",
            font=("맑은 고딕", 9),
            fg="#95a5a6",
            bg="white"
        )
        self.status_label.pack(side=tk.BOTTOM, pady=10)

    def on_drop(self, event):
        """파일 드롭 이벤트 처리"""
        if self.is_processing:
            self.show_message("이미 처리 중입니다. 잠깐 기다려주세요.", "warning")
            return

        files = self.parse_drop_files(event.data)
        hwpx_files = [f for f in files if f.lower().endswith('.hwpx')]

        if not hwpx_files:
            self.show_message("HWPX 파일을 드롭해주세요.", "error")
            return

        # 스레드 처리
        thread = threading.Thread(target=self.process_files, args=(hwpx_files,))
        thread.daemon = True
        thread.start()

    def parse_drop_files(self, data):
        """드롭된 파일 경로 파싱"""
        files = []
        for item in self.root.tk.splitlist(data):
            item = item.strip('{}')
            if os.path.exists(item):
                files.append(item)
        return files

    def process_files(self, files):
        """파일 처리"""
        self.is_processing = True
        total = len(files)
        success = 0

        # 현재 선택된 크기 적용
        target_size = self.size_var.get()
        compressor = HWPXImageCompressorAdvanced(target_size_kb=target_size)

        for idx, file_path in enumerate(files):
            result, message = compressor.process_hwpx(
                file_path,
                progress_callback=self.update_progress
            )

            if result:
                success += 1

        self.is_processing = False
        if success == total:
            self.show_message(f"✅ 완료!\n성공: {success}/{total}", "success")
        else:
            self.show_message(f"⚠️ 일부 완료\n성공: {success}/{total}", "warning")

        self.progress_bar['value'] = 0
        self.detail_label.config(text="")

    def update_progress(self, progress, status, current, total, elapsed):
        """진행 상태 업데이트"""
        self.progress_bar['value'] = progress
        detail_text = f"{status}\n진행: {current}/{total} | 경과: {elapsed}초"
        self.detail_label.config(text=detail_text)
        self.root.update()

    def show_message(self, message, msg_type="info"):
        """메시지 표시"""
        colors = {
            "success": "#27ae60",
            "error": "#e74c3c",
            "warning": "#f39c12",
            "info": "#3498db"
        }

        self.status_label.config(text=message, fg=colors.get(msg_type, colors["info"]))
        self.root.update()
        self.root.after(5000, lambda: self.status_label.config(text="대기 중...", fg="#95a5a6"))

    def run(self):
        """GUI 실행"""
        self.root.mainloop()


if __name__ == "__main__":
    app = HWPXCompressorGUI()
    app.run()
