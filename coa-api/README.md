# COA API

사내 ERP Firebird DB를 SELECT 전용으로 조회하는 Node.js API입니다.

## 로컬 준비

1. Node.js 20 이상을 설치합니다.
2. `.env.example`을 `.env`로 복사하고 실제 값을 입력합니다.
3. `npm install`을 실행합니다.
4. `npm test` 후 `npm start`를 실행합니다.

API 호출에는 `Authorization: Bearer <API_KEY>` 헤더가 필요합니다. `/` 상태 확인만 인증 없이 사용할 수 있습니다.

## Snapshot

기본 저장 위치는 `data/snapshots`입니다. 운영 서버의 기존 폴더를 계속 사용하려면 `.env`에 다음처럼 지정할 수 있습니다.

```text
SNAPSHOT_DIR=shipment_snapshots
```

기존 JSON을 이동하기 전에는 운영 서버에서 API 프로세스를 중지하고 백업해야 합니다.

## 변경사항 처리 이력

`PRODUCTION_PLAN`과 `SHIPMENT_DISPOSAL`은 서로 독립적으로 처리 완료 상태를 관리합니다.
처리 이력은 기본적으로 `data/change_acknowledgements.ndjson`에 저장되며 ERP DB에는 기록하지 않습니다.
