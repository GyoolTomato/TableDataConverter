# TableDataConverter

Excel로 관리하는 게임 테이블을 Unity용 데이터와 C# 코드로 변환하고, 선택한 테이블을 SQLite DB에 반영하는 Windows WinForms 도구입니다.

## 주요 기능

- 실행 폴더의 `_*.xlsx` 파일 자동 검색
- Excel이 열려 있어도 마지막으로 저장된 내용을 메모리 스냅샷으로 읽기
- JSON 형식의 `.bytes` 데이터 생성
- 테이블 C# 클래스, enum 및 `TableDataLoader.cs` 생성
- `.bytes`와 `.cs` 출력 폴더 개별 선택 및 탐색기 열기
- SQLite `.db` 파일 선택
- 체크한 Excel 테이블만 SQLite DB에 반영
- DB 반영 전 자동 백업 및 트랜잭션 처리
- 선택한 출력 경로와 DB 경로를 외부 설정 파일에 저장

## 실행 환경

- Windows x64
- .NET 8 Desktop Runtime x64

소스 빌드에는 .NET 8 SDK가 필요합니다.

## Excel 파일 규칙

Excel 파일과 `TableDataConverter.exe`를 같은 작업 폴더에 둡니다. 파일명은 다음 형식을 사용합니다.

```text
_<3자리 테이블 번호>_<테이블 이름>.xlsx
```

예시:

```text
_000_Global_Enum.xlsx
_101_Items.xlsx
_104_Missions.xlsx
_900_CommonText.xlsx
```

Excel 데이터는 다음 행 구조를 사용합니다.

| 행 | 용도 |
| --- | --- |
| 1행 | 설명 등 자유 영역 |
| 2행 | 변수명 또는 DB 열 이름 |
| 3행 | 자료형 |
| 4행 이후 | 실제 데이터 |

`.`으로 시작하는 열과 이름이 없는 열은 변환 대상에서 제외됩니다.

### 지원 타입

| Excel 타입 | JSON/C# 처리 | SQLite 타입 |
| --- | --- | --- |
| `byte`, `short`, `int`, `long` | 정수 | `INTEGER` |
| `float`, `double`, `decimal` | 실수 | `REAL` |
| `bool` | Boolean | `INTEGER` (`0`/`1`) |
| `string` | 문자열 | `TEXT` |
| enum 등 사용자 타입 | 문자열 | `TEXT` |

### 반복 배열

2행에서 같은 배열 이름을 대괄호로 반복해 선언할 수 있습니다.

```text
[rewardKeys] | [rewardKeys] | [rewardKeys]
```

- `.bytes`에서는 하나의 JSON 배열 `rewardKeys`로 생성됩니다.
- SQLite에서는 `rewardKeys_0`, `rewardKeys_1`, `rewardKeys_2`로 생성됩니다.
- 배열의 첫 열에 타입을 선언하면 뒤의 같은 배열 열은 해당 타입을 이어받습니다.

## Data 변환

1. `.bytes 저장 경로`에서 출력 폴더를 선택합니다.
2. `.cs 저장 경로`에서 출력 폴더를 선택합니다.
3. 필요한 경우 `목록 새로고침`을 누릅니다.
4. `전체 Data 변환`을 누릅니다.

Data 변환은 목록의 체크 상태와 관계없이 검색된 Excel 전체를 대상으로 합니다. 출력 폴더의 기존 파일을 자동으로 삭제하지 않으며, 같은 이름의 생성 파일만 덮어씁니다.

## SQLite DB 반영

1. `SQLite DB 파일`에서 대상 `.db`를 선택합니다.
2. 목록에서 반영할 테이블을 체크합니다.
3. `선택한 테이블 DB 반영`을 누릅니다.
4. 확인 창에서 대상 개수를 확인하고 반영합니다.

DB 테이블 이름은 Excel 확장자를 제외한 파일명을 그대로 사용합니다.

```text
_101_Items.xlsx -> _101_Items
```

DB 반영 대상 Excel에는 2행에 `key` 열이 반드시 있어야 하며, 이 열은 SQLite의 `PRIMARY KEY`가 됩니다. 빈 key와 중복 key가 있으면 반영을 중단합니다.

### 선택 제한

테이블 번호의 첫 자리가 `0` 또는 `9`인 Excel은 DB 반영 대상으로 선택할 수 없습니다. 목록에는 표시되지만 체크박스가 비활성화됩니다.

| 파일 | DB 선택 |
| --- | --- |
| `_000_Global_Enum.xlsx` | 불가 |
| `_101_Items.xlsx` | 가능 |
| `_800_Notice.xlsx` | 가능 |
| `_900_CommonText.xlsx` | 불가 |
| `_999_SystemText.xlsx` | 불가 |

### 반영 방식과 주의 사항

- 선택한 테이블은 기존 테이블을 삭제하고 Excel 스키마와 데이터로 다시 생성합니다.
- Excel에서 제거된 행과 열도 DB에서 제거됩니다.
- 선택한 모든 테이블은 하나의 SQLite 트랜잭션으로 처리됩니다.
- 실패하면 트랜잭션을 롤백합니다.
- 반영 직전에 대상 DB와 같은 폴더에 백업을 생성합니다.
- 재생성되는 테이블에 별도 인덱스나 트리거가 있었다면 함께 제거되므로 마스터 데이터 테이블에만 사용해야 합니다.

백업 파일 예시:

```text
dev_20260903_173159.backup.db
```

## 설정 파일

경로 선택값은 실행 파일에 포함하지 않고 실행 파일 옆의 파일에 저장합니다.

```text
TableDataConverter.settings.json
```

저장되는 항목:

- `.bytes` 출력 경로
- `.cs` 출력 경로
- SQLite DB 파일 경로

프로그램을 다시 실행하면 해당 설정을 자동으로 불러옵니다. 다른 PC로 설정까지 옮기려면 EXE와 설정 파일을 함께 복사합니다.

## 개발 및 빌드

```powershell
dotnet restore TableDataConverter.sln --runtime win-x64
dotnet build TableDataConverter.sln --no-restore
```

주요 의존성:

- [ClosedXML](https://github.com/ClosedXML/ClosedXML): Excel 읽기
- [Microsoft.Data.Sqlite](https://learn.microsoft.com/dotnet/standard/data/sqlite/): SQLite 반영 및 백업

## Windows x64 단일 파일 게시

.NET 런타임을 포함하지 않는 Release 단일 EXE는 다음 명령으로 게시합니다.

```powershell
dotnet publish TableDataConverter/TableDataConverter.csproj `
  -c Release `
  -r win-x64 `
  --self-contained false `
  -p:PublishSingleFile=true `
  -p:DebugType=None `
  -p:DebugSymbols=false `
  -o dist/win-x64
```

게시 결과:

```text
dist/win-x64/TableDataConverter.exe
```

이 배포본에는 .NET 런타임이 포함되지 않으므로 실행 PC에 .NET 8 Desktop Runtime x64가 설치되어 있어야 합니다.
