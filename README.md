# Table Data Converter

Excel 테이블 데이터를 Unity 프로젝트에서 사용할 수 있는 C# 스크립트와 `.bytes` 데이터로 변환하는 WinForms 도구입니다.

## 주요 기능

- 현재 실행 폴더의 `_*.xlsx` 파일 목록 표시
- Excel 테이블을 JSON 형식의 `.bytes` 파일로 변환
- 테이블별 C# 데이터 클래스 자동 생성
- enum C# 파일 자동 생성
- `TableDataLoader.cs` 자동 생성

## 사용 환경

- Windows
- .NET 8
- Visual Studio 또는 `dotnet` CLI
- ClosedXML

## 입력 파일 규칙

- 변환 대상 파일은 프로그램 실행 폴더에 위치해야 합니다.
- 파일명은 `_`로 시작하고 `.xlsx` 확장자를 가져야 합니다.
- 파일명의 두 번째 문자가 `0`이면 enum 테이블로 처리됩니다.
  - 예: `_0ItemType.xlsx`
- 그 외 파일은 일반 데이터 테이블로 처리됩니다.
  - 예: `_Item.xlsx`

## Excel 작성 규칙

| 행 | 내용 |
| --- | --- |
| 2행 | 변수명 |
| 3행 | 타입 |
| 4행 이후 | 실제 데이터 |

지원 타입:

- `int`
- `long`
- `double`
- `float`
- `string`
- enum 타입

## 출력 결과

변환 시 Unity 프로젝트 기준으로 다음 파일들이 생성됩니다.

| 출력 위치 | 내용 |
| --- | --- |
| `Assets/Scripts/_Common/GlobalData` | enum C# 파일 |
| `Assets/Scripts/_Common/Tables` | 테이블 C# 클래스, `TableDataLoader.cs` |
| `Assets/Tables` | 변환된 `.bytes` 데이터 |

## 사용 방법

1. 변환할 `_*.xlsx` 파일을 프로그램 실행 폴더에 배치합니다.
2. 프로그램을 실행합니다.
3. `Refresh` 버튼으로 파일 목록을 갱신합니다.
4. `Confirm` 버튼으로 변환을 실행합니다.

## 게시 방법

Visual Studio 게시 기능 대신 CLI로 게시합니다.

```bat
cd C:\GitHub\TableDataConverter
dotnet publish -c Release -r win-x64 --self-contained false /p:PublishSingleFile=true
```

또는 `CLI.bat`을 실행합니다.

## 연동 프로젝트

게시된 프로그램은 Unity 테이블 프로젝트에서 사용합니다.

- [UnityPortfolioProjectTables](https://github.com/GyoolTomato/UnityPortfolioProjectTables)

## 변경 내역

### v1.0

- 버전 네이밍
- 테이블 데이터 변환 기능
- `TableDataLoader.cs` 자동 생성
