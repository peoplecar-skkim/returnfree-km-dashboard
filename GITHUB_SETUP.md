# returnfree-km-dashboard 배포 설정 (최초 1회만 하면 됨)

## 1. GitHub에서 저장소 생성
1. github.com 로그인 → 우측 상단 "+" → New repository
2. Repository name: `returnfree-km-dashboard`
3. **Public** 선택 (GitHub Pages 무료 티어는 Public 저장소만 지원)
4. 나머지 기본값으로 "Create repository"

## 2. GitHub Pages 설정
1. 방금 만든 저장소 → **Settings > Pages**
2. Source: **Deploy from a branch**
3. Branch: **main**, 폴더: **/docs** 선택 → Save
4. 배포되면 아래 형태의 URL이 뜸 (몇 분 정도 걸릴 수 있음):
   `https://<깃허브아이디>.github.io/returnfree-km-dashboard/`

## 3. 로컬 폴더를 이 저장소와 연결 (최초 1회)

터미널(명령 프롬프트/PowerShell/Git Bash)에서, 이 `turucar_dashboard` 폴더 안으로 이동한 뒤:

```
git init
git add .
git commit -m "최초 커밋"
git branch -M main
git remote add origin https://github.com/<깃허브아이디>/returnfree-km-dashboard.git
git push -u origin main
```

- `<깃허브아이디>` 자리에 실제 GitHub 아이디를 넣으세요.
- push 할 때 GitHub 로그인을 요구하면, 비밀번호 대신 **Personal Access Token(PAT)**을 입력해야 할 수도 있어요 (GitHub이 2021년부터 비밀번호 직접 입력을 막았어요). PAT은 GitHub 우측상단 프로필 → Settings → Developer settings → Personal access tokens에서 발급받으면 됩니다. (또는 SSH 키를 등록해두면 그걸로 인증됨)

## 4. 이후로는 그냥 `update_dashboard.bat` 더블클릭만 하면 됨

이 배치파일이 하는 일:
1. `generate_data.py` 실행 → 최신 데이터로 `output/index.html` 갱신
2. `output/index.html`을 `docs/index.html`로 복사
3. `git add / commit / push` 자동 실행 → GitHub Pages에 반영 (1~2분 후 URL에서 확인 가능)
4. 로컬에서도 바로 볼 수 있게 index.html을 엽니다

`.gitignore`에 `data/`와 `output/`이 들어있어서, 원본 데이터 파일이나 작업 중간 파일은 GitHub에 절대 올라가지 않고 `docs/index.html`(최종 배포본, 이미 데이터가 다 박혀있는 단독 실행 파일)만 올라갑니다.

## 참고
- push가 실패해도(네트워크 문제, 인증 만료 등) 로컬 `output/index.html`은 정상적으로 갱신되니 당장 급하면 그 파일을 열어서 보면 됩니다.
- 매주 파일(첫이용 마스터, 스테이션등록, 주유충전비 등)만 최신본으로 교체해두면, bat 실행 한 번으로 데이터 갱신 + 배포까지 끝납니다.
