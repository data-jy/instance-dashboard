# 인스턴스 리소스 분석 현황판 (Instance Dashboard)

**포트**: 5001
**기술 스택**: Python HTTP Server, Pandas, OpenPyXL
**목적**: NHN Cloud 인스턴스 리소스 사용률을 분석하고 최적화된 사양을 제안하는 웹 기반 대시보드

---

## 🎯 주요 기능

- **리소스 분석**
  - 📊 CPU/메모리 사용률 실시간 분석
  - 🎯 목표 사용률 대비 비교 (기본: CPU 70%, MEM 80%)
  - 📈 사용률 추이 시각화

- **사양 최적화**
  - 🔍 NHN Cloud 인스턴스 사양 매핑
  - 💡 최적화된 사양 자동 추천
  - 📋 권장사항 상세 분석

- **보고서 생성**
  - 📄 Excel 분석 보고서 자동 생성
  - 📑 상세 시트별 분류
  - 💾 타임스탬프로 버전 관리

- **웹 UI**
  - 🌐 대시보드 뷰
  - 📊 상세 분석 페이지
  - 📥 파일 업로드 & 분석

---

## 🚀 빠른 시작

### 사전 요구사항
- Python 3.8+
- pandas
- openpyxl

### 설치

```bash
# 저장소 클론
git clone https://github.com/data-jy/instance-dashboard.git
cd instance-dashboard

# 필요한 패키지 자동 설치
python server.py
# (처음 실행 시 자동으로 필요한 패키지 설치)
```

### 실행

```bash
# 기본 포트 (5001)에서 실행
python server.py

# 또는 다른 포트 지정
python server.py --port 8000

# 자동 브라우저 열기 포함
python server.py --auto-open
```

**접속**: http://localhost:5001

---

## 📁 프로젝트 구조

```
instance-dashboard/
├── server.py              # HTTP 서버 & 웹 API
├── analyze.py             # 리소스 분석 엔진
├── index.html             # 메인 대시보드 UI
├── resource_analyze.html  # 상세 분석 페이지
└── (업로드된 데이터 폴더)
```

---

## 💻 사용 방법

### 방법 1: 웹 UI (권장)

1. **브라우저 접속**: http://localhost:5001
2. **Pod 데이터 파일 업로드** (YAML 또는 JSON)
3. **사양 정보 파일 업로드** (Excel 파일)
4. **분석 시작** 버튼 클릭
5. **결과 보고서 다운로드**

### 방법 2: 명령어 라인

```bash
# 데이터 폴더와 사양 파일로 분석
python analyze.py ./data 인망_행망_인스턴스자원현황.xlsx

# 커스텀 목표 사용률 설정
python analyze.py ./data 사양.xlsx --cpu-target 65 --mem-target 75

# 출력 파일명 지정
python analyze.py ./data 사양.xlsx --out 분석결과.xlsx
```

---

## 📊 분석 기능

### CPU/메모리 등급

| 등급 | CPU 사용률 | 메모리 사용률 | 설명 |
|------|-----------|-------------|------|
| 🟢 좋음 | < 70% | < 80% | 최적 상태 |
| 🟡 주의 | 70-90% | 80-95% | 모니터링 필요 |
| 🔴 경고 | > 90% | > 95% | 즉시 조치 필요 |

### NHN Cloud 인스턴스 타입

**m2 계열** (vCPU:RAM = 1:2)
- 2vCPU/4GB, 4vCPU/8GB, 8vCPU/16GB, 16vCPU/32GB, 32vCPU/64GB

**c2 계열** (vCPU:RAM = 1:1)
- 2vCPU/2GB, 4vCPU/4GB, 8vCPU/8GB, 16vCPU/16GB

**r2 계열** (vCPU:RAM = 1:4~8)
- 2vCPU/8GB, 4vCPU/16GB, 8vCPU/32GB, 8vCPU/64GB

---

## 📋 Excel 보고서 구조

자동 생성되는 Excel 파일에는 다음 시트가 포함됩니다:

### 1. **요약 (Summary)**
- 전체 분석 통계
- 카테고리별 분류
- 주요 권장사항

### 2. **상세 분석 (Detail)**
- 각 인스턴스별 상세 정보
- 현재 사양 vs 권장 사양
- 절약 가능 비용 계산

### 3. **Pod 정보 (Pods)**
- Kubernetes Pod 정보
- Namespace별 분류
- 리소스 요청/한계 값

### 4. **Pod 요약 (Pod Summary)**
- Pod별 종합 통계
- 리소스 배분 현황

---

## 🔧 API 엔드포인트

| 메서드 | 경로 | 설명 |
|--------|------|------|
| GET | `/` | 메인 대시보드 |
| GET | `/analyze` | 분석 페이지 |
| POST | `/upload` | 파일 업로드 & 분석 |
| GET | `/download/<filename>` | 결과 보고서 다운로드 |

### 파일 업로드

```bash
# 멀티파트 폼 데이터로 전송
curl -X POST http://localhost:5001/upload \
  -F "pod_file=@pod_data.yaml" \
  -F "spec_file=@specs.xlsx"
```

---

## 📊 데이터 포맷

### Pod 데이터 파일 (YAML/JSON)
```yaml
apiVersion: v1
kind: Pod
metadata:
  name: my-pod
  namespace: default
spec:
  containers:
  - name: app
    image: myapp:latest
    resources:
      requests:
        cpu: 100m
        memory: 128Mi
      limits:
        cpu: 500m
        memory: 512Mi
```

### 사양 정보 파일 (Excel)
| 인스턴스명 | vCPU | RAM(GB) | 용도 | 현재비용 |
|-----------|------|--------|------|---------|
| prod-01 | 4 | 8 | 웹서버 | 100,000 |
| prod-02 | 8 | 16 | DB | 200,000 |

---

## 🎨 UI 기능

### 대시보드 뷰
- 📊 실시간 리소스 모니터링
- 🎯 권장사항 우선순위 표시
- 💰 비용 절감 예상액

### 분석 페이지
- 📈 사용률 추이 그래프
- 📋 상세 데이터 테이블
- 🔍 필터링 & 검색

---

## 🛠️ 설정 옵션

### 커맨드라인 옵션

```bash
python server.py [옵션]

옵션:
  --port PORT              실행 포트 (기본: 5001)
  --cpu-target PCT        CPU 목표 사용률 % (기본: 70)
  --mem-target PCT        메모리 목표 사용률 % (기본: 80)
  --out FILENAME          출력 파일명 (기본: 리소스_분석보고서_YYYYMMDD_HHMM.xlsx)
  --auto-open             자동으로 브라우저 열기
```

### 환경변수

```bash
export INSTANCE_DASHBOARD_PORT=5001
export CPU_TARGET=65
export MEM_TARGET=75
python server.py
```

---

## 📞 문제 해결

**패키지 설치 오류**
```bash
# 수동 설치
pip install pandas openpyxl --break-system-packages
```

**포트 충돌**
```bash
# 다른 포트에서 실행
python server.py --port 8001
```

**파일 업로드 오류**
- 파일 형식 확인 (YAML, JSON, Excel)
- 파일 크기 제한 확인

**분석 결과 오류**
- 데이터 포맷이 스펙에 맞는지 확인
- 필수 필드 누락 여부 확인

---

## 📚 참고

- NHN Cloud 인스턴스 사양: https://docs.nhncloud.com/
- Kubernetes 리소스 요청/한계: https://kubernetes.io/docs/

---

## 📝 라이선스

내부 분석 도구

---

## 👥 지원

문제 발생 시 GitHub Issues에 보고해주세요.
https://github.com/data-jy/instance-dashboard/issues

---

**최종 수정**: 2026-04-03
**버전**: 1.0.0
