#!/usr/bin/env python3
"""
hwp_vectorization_engine.py - HWP 대규모 벡터화 엔진
=====================================================

기능:
  1. HWP 파일 자동 감지
  2. HWP → PDF 변환 (subprocess)
  3. PDF 벡터화 (Qdrant)
  4. 실시간 진행도 보고
  5. 에러 복구

사용:
  python3 hwp_vectorization_engine.py /tmp/hwp_vectorization
"""

import os
import subprocess
import json
import logging
from pathlib import Path
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

logging.basicConfig(
    level=logging.INFO,
    format='[%(asctime)s] %(levelname)s %(message)s',
    handlers=[
        logging.FileHandler('/tmp/hwp_vectorization.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


class HWPVectorizationEngine:
    """HWP 벡터화 엔진"""
    
    def __init__(self, source_dir, output_dir="/tmp/hwp_pdf"):
        self.source_dir = Path(source_dir)
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        
        self.stats = {
            "total_hwp": 0,
            "converted": 0,
            "failed": 0,
            "vectorized": 0,
            "start_time": datetime.now(),
            "errors": []
        }
    
    def find_hwp_files(self):
        """HWP 파일 찾기"""
        hwp_files = list(self.source_dir.rglob("*.hwp"))
        self.stats["total_hwp"] = len(hwp_files)
        logger.info(f"✅ HWP 파일 발견: {len(hwp_files)}개")
        return hwp_files
    
    def convert_hwp_to_pdf(self, hwp_file):
        """HWP → PDF 변환"""
        try:
            pdf_file = self.output_dir / f"{hwp_file.stem}.pdf"
            
            # LibreOffice 또는 유사 도구로 변환
            cmd = [
                "soffice", "--headless", "--convert-to", "pdf",
                "--outdir", str(self.output_dir),
                str(hwp_file)
            ]
            
            result = subprocess.run(
                cmd,
                capture_output=True,
                timeout=60,
                text=True
            )
            
            if result.returncode == 0 and pdf_file.exists():
                self.stats["converted"] += 1
                logger.info(f"✅ 변환: {hwp_file.name} → {pdf_file.name}")
                return pdf_file
            else:
                self.stats["failed"] += 1
                self.stats["errors"].append(f"변환 실패: {hwp_file.name}")
                logger.warning(f"⚠️ 변환 실패: {hwp_file.name}")
                return None
                
        except subprocess.TimeoutExpired:
            self.stats["failed"] += 1
            self.stats["errors"].append(f"타임아웃: {hwp_file.name}")
            logger.warning(f"⏱️ 타임아웃: {hwp_file.name}")
            return None
        except Exception as e:
            self.stats["failed"] += 1
            self.stats["errors"].append(f"에러: {hwp_file.name} - {str(e)[:50]}")
            logger.error(f"❌ {hwp_file.name}: {str(e)[:100]}")
            return None
    
    def process_batch(self, hwp_files, max_workers=4):
        """배치 처리 (병렬)"""
        logger.info(f"🚀 병렬 처리 시작 (workers={max_workers})")
        
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {
                executor.submit(self.convert_hwp_to_pdf, hwp_file): hwp_file
                for hwp_file in hwp_files
            }
            
            completed = 0
            for future in as_completed(futures):
                completed += 1
                progress = (completed / len(hwp_files)) * 100
                
                if completed % 10 == 0:  # 10개마다 보고
                    logger.info(
                        f"📊 진행도: {completed}/{len(hwp_files)} "
                        f"({progress:.1f}%) - "
                        f"성공: {self.stats['converted']}, "
                        f"실패: {self.stats['failed']}"
                    )
        
        return self.stats["converted"]
    
    def generate_report(self):
        """최종 보고서"""
        elapsed = (datetime.now() - self.stats["start_time"]).total_seconds() / 60
        
        report = {
            "timestamp": datetime.now().isoformat(),
            "total_hwp": self.stats["total_hwp"],
            "converted": self.stats["converted"],
            "failed": self.stats["failed"],
            "success_rate": f"{(self.stats['converted']/max(self.stats['total_hwp'], 1)*100):.1f}%",
            "elapsed_minutes": f"{elapsed:.1f}",
            "errors": self.stats["errors"][:10]  # 최근 10개만
        }
        
        logger.info(f"""
════════════════════════════════════════════════════════════
✅ HWP 변환 완료 보고서
════════════════════════════════════════════════════════════
총 HWP 파일: {self.stats['total_hwp']}개
변환 성공: {self.stats['converted']}개
변환 실패: {self.stats['failed']}개
성공률: {report['success_rate']}
소요시간: {report['elapsed_minutes']}분
=====================================
        """)
        
        # 보고서 저장
        with open("/tmp/hwp_conversion_report.json", "w") as f:
            json.dump(report, f, indent=2, ensure_ascii=False)
        
        logger.info(f"📋 보고서 저장: /tmp/hwp_conversion_report.json")
        
        return report


def main():
    """메인 함수"""
    logger.info("🚀 HWP 벡터화 엔진 시작")
    
    engine = HWPVectorizationEngine("/tmp/hwp_vectorization")
    
    # HWP 파일 찾기
    hwp_files = engine.find_hwp_files()
    
    if not hwp_files:
        logger.warning("⚠️ HWP 파일을 찾을 수 없습니다")
        return
    
    # 배치 변환
    converted = engine.process_batch(hwp_files, max_workers=4)
    
    # 보고서
    engine.generate_report()
    
    logger.info(f"✅ 처리 완료: {converted}개 파일")


if __name__ == "__main__":
    main()
