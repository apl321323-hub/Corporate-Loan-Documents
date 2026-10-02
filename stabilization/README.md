# 데이터 안정화 검증

이 폴더의 테스트는 운영 Supabase나 애플리케이션 서버를 호출하지 않는다. 공통 업로드 쓰기 범위, 빈 재무 데이터 덮어쓰기 방지, Supabase 일괄 읽기·쓰기 요청 수, 계약·결산·취급건수 업로드의 `asset_data` 교차 저장 차단을 오프라인으로 검사한다.

```powershell
python -B -m unittest discover -s stabilization -p "test_*.py" -v
```

현재 자산 화면의 계약·결산·취급건수 값은 원본 키를 보존한 채 `/api/data/asset` 응답 복사본에서 파생한다. 운영 데이터 변경과 배포는 별도 검증 후 수행한다.
