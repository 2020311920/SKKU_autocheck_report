# SKKU_autocheck_report

**본 프로젝트는 성균관대학교 SW중심대학사업단(이하 SKKU소중사업단)에서 진행하고 있는**
**"성대의 성대한 특강"(이하 성성특)에 대한 보고서 자동화를 위해 시작되었습니다.**

*이는 SKKU소중사업단 가치확산 서포터즈로 활동하고 있는 강성철의 개인 프로젝트입니다.</br>*

정보를 추출하여
1. 보고서를 작성한 학생이 누군지
2. 수강한 강좌를 이미 이수한 이력이 있는지
3. 있다면 언제, 없다면 새롭게 DB에 추가
4. 관련 정보를 사용자에게 출력
5. (추가 편의기능)

이를 통해서 보다 편리하게 업무를 진행하도록 하는 것이 본 프로젝트의 목표입니다.

- 단기적인 목표는 Excel에서의 완전한 자동화</br>
- 장기적으로는 DB(MYSQL)를 활용하여 웹 서비스를 제공하고자 합니다.

해당 프로젝트는 Excel내의 VBA와 Python을 활용하여 진행되었습니다.
추후에 MYSQL과 JS등을 활용하여 웹서비스로 제공될 예정입니다.

해당 프로젝트는 24년 2월부터 진행되었으며, 최종적으로는 24년 2분기 안으로 제작하는 것이 목표입니다.

## 현재 기능 추가 및 버그 수정 등 패치 이력

## ver3.1
- **파이썬 절대 경로 및 코드 경로 수정**:
  - **버그 수정**:
    - 이전까지는 현재 작업 중인 로컬 상의 절대 경로를 사용하여 타인이 사용할 수 없었습니다. 이를 수정하여, 파일을 어디에 다운받든 편하게 사용할 수 있습니다.

- **사전 패키지 설치**:
  - **용이성 향상**:
    - 파일을 처음 받고도 파이썬만 설치되어있다면 사전패키지 설치 파일로 쉽게 모든 패키지를 다운 받을 수 있습니다.
      
- **타 이용자 서비스 가능**:
  - **6/21부로 타 이용자에게 서비스 제공**:
    - 앞으로의 피드백을 통해 향상 예정


## ver2.5.1
- **보고서 내 강의 중복 알고리즘 오류 수정**:
  - **버그 수정**:
    - 이전 버전에서 발생한 보고서 내 강의 중복을 처리하는 알고리즘의 오류를 수정했습니다. 이제 올바르게 중복을 식별하고 처리할 수 있습니다.

- **분산된 코드 작업**:
  - **코드 리팩토링**:
    - 프로젝트의 코드를 보다 모듈화하고 구조화하여 유지보수와 확장성을 향상시켰습니다. 이제 코드의 가독성이 좋아졌고, 새로운 기능 추가 및 버그 수정이 보다 용이해졌습니다.

- **각 학생별 총 이수강의 수 추출 기능 추가**:
  - **통계 기능 추가**:
    - 각 학생이 이수한 총 강의 수를 추출하여 통계 데이터로 제공하는 기능이 추가되었습니다. 이를 통해 학생들의 학습 현황을 더 잘 이해할 수 있습니다.
      <img width="989" alt="image" src="https://github.com/2020311920/SKKU_autocheck_report/assets/80453145/6eb59b90-5864-4579-86bc-4a674e84c230">

- **전체 DB내의 중복된 데이터 자동 제거기능 추가**:
  - **데이터 정리 기능 추가**:
    - 전체 데이터베이스(DB) 내에서 중복된 항목을 자동으로 식별하고 제거하는 기능이 추가되었습니다. 이를 통해 데이터베이스의 정확성과 효율성을 유지할 수 있습니다.
      ![image](https://github.com/2020311920/SKKU_autocheck_report/assets/80453145/9669c6a1-9da8-4c3b-a44f-0fcb9860c0bd)

- **엑셀 내 VBA코드 및 시트 1차 보안기능 추가**:
  - **보안 강화**:
    - 엑셀 파일 내에 사용된 VBA 코드와 시트에 보안 기능을 추가하여 외부로부터의 불법적인 접근을 방지하고 데이터의 안전성을 보호하는 데 도움이 됩니다.

## ver2.5
- **파일 형식 다양화**:
  - **pdf 파일 지원 추가** : 이제는 PDF 파일도 지원합니다. 사용자가 편리하게 보고서를 업로드하여 정보를 추출할 수 있습니다.
  
- **미이수 강의 대해 한번에 일괄적으로 한번에 반영하는 기능 추가 & 추가 인정될 시간 시각적 표현 추가**:
  - **편의성 개선**:
    - 학생들이 이수하지 않은 강의를 한 번에 일괄적으로 데이터에 반영할 수 있는 기능이 추가되었습니다. 이를 통해 보다 효율적인 데이터 관리가 가능해졌습니다.
      ![image](https://github.com/2020311920/SKKU_autocheck_report/assets/80453145/e3405cc9-0ba4-437d-a891-1743205a3759)
      ![image](https://github.com/2020311920/SKKU_autocheck_report/assets/80453145/b9f08be0-572e-4738-9660-724b153728b1)

- **보고서에서 강의를 중복할 경우 이를 시각적으로 보여주는 기능 추가**:
  - **사용자 경험 개선**:
    - 보고서 작성 시 강의를 중복한 경우에 대해 시각적으로 경고를 표시하여 사용자가 실수를 방지할 수 있도록 하는 기능이 추가되었습니다.
      ![image](https://github.com/2020311920/SKKU_autocheck_report/assets/80453145/b98ee978-aa1b-4971-b5cf-668c5854cb9f)


## ver2.4
- **파일 형식 다양화**:
  - **doc 파일 지원 추가** : 이제는 DOC 파일도 지원합니다. 사용자가 편리하게 보고서를 업로드하여 정보를 추출할 수 있습니다.

## ver2.3
- **파일 형식 다양화**:
  - **docx 파일 지원 추가** : 이전에는 hwp 형식만을 지원했지만 이제는 docx 파일도 지원합니다. 사용자가 편리하게 보고서를 업로드하여 정보를 추출할 수 있습니다.

## ver2.2
- **영상 고유 코드 "유효도 검사" 기능 추가**:
  - **기능 설명**:
    - 영상 고유 코드를 악의적으로 지어낼 경우를 대비하여, 해당 코드가 유효한지 판단하는 기능을 추가하였습니다.
    - <img width="680" alt="image" src="https://github.com/2020311920/SKKU_autocheck_report/assets/80453145/72868ba0-e11b-43d2-8bf4-528bfe26609b">


## ver2.1
- **중복 검사 정확도 향상**:
  - **영상별 고유코드로 구분하는 기능 추가**
  - **설명**:
    - 영상 고유 코드를 통해 각 영상을 식별하고, 이를 활용하여 중복을 검사하는 기능이 추가되었습니다. 보고서 양식 상 강의명이 달라지는 문제를 해결하기 위해 도입되었습니다.
    - 이를 통해 사용자는 학생별 보고서에 따라 같은 강좌명임에도 서로 미세하게 다른 경우에 영상 고유 코드를 통해 정확한 강의를 식별할 수 있게 되었습니다.
    - ![이미지](https://github.com/2020311920/SKKU_autocheck_report/assets/80453145/f2efde66-5d05-44f6-ad64-588daa1fc2cf)

## ver1.3
- **기능 추가**:
  - **더 쉬워진 데이터 반영** : 이수하지 않은 경우, 데이터에 옮겨질 형식을 미리보여주며, 버튼을 누를경우 자동으로 데이터에 반영됩니다.
  - ![이미지](https://github.com/2020311920/SKKU_autocheck_report/assets/80453145/23edb0d8-c6e8-4110-bd25-02a3864175e5)

## ver1.2
- **기능 개선**:
  - **파일 업로드 기능 추가** : 사용자가 이름, 강의명을 입력하는 대신 파일을 업로드하여 자동으로 보여지도록 기능 추가 (hwp 형식만 지원)
  - **동시 일치 항목 추가** : 동시 일치 항목을 2개에서 4개로 늘려 정확도 향상 (학생명, 강좌명 -> 학생명, 강좌명, 학번, 학과)
  - **성능 개선** : 동시 일치 여부 함수를 Excel내의 함수에서 VBA코드로 변경하여 정확도 및 속도 향상 (VBA코드 내에서 python코드 실행)
  - ![이미지](https://github.com/2020311920/SKKU_autocheck_report/assets/80453145/88ef6c66-a3c5-4b95-80e3-10a6bf9b176a)

## ver1.1
- **기능 추가**:
  - **학생 이름, 강좌명 동시 일치 여부 기능**
  - ![이미지](https://github.com/2020311920/SKKU_autocheck_report/assets/80453145/bb4c0835-270c-4ce0-9719-ee24bd155644)










 
  


