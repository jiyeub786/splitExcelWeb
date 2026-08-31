---
version: alpha
name: splitExcelWeb-light-theme
description: splitExcelWeb의 다크/라이트 토글 기능을 위한 라이트 테마 토큰 문서. design.md(Binance 스타일 분석)의
  브랜드 액센트·타이포그래피·반경 토큰은 그대로 공유하고, 캔버스/표면/헤어라인/텍스트 톤만 라이트 모드용으로
  다시 정의한다. design.md가 "이 브랜드의 색이 무엇인가"를 정의한다면, design2.md는 "그 브랜드를 밝은
  배경에서 어떻게 보여줄 것인가"를 정의한다.
---

## 개요

`splitExcelWeb`은 기본이 다크 테마(design.md 기준)이며, 헤더의 🌙/☀️ 토글 버튼으로 라이트 테마를
전환할 수 있다. 라이트 테마는 완전히 새로운 브랜드가 아니라 **같은 디자인 시스템의 밝은 배경 변형**이다 —
design.md가 이미 문서화한 Binance의 "라이트 모드 (거래성)" 섹션의 토큰(`canvas-light`,
`surface-strong-light`, `hairline-on-light`, `border-strong`, `ink` 등)을 그대로 재사용한다.

**공유되는 것 (테마와 무관하게 고정)**:
- 브랜드 액센트 `{colors.primary}`(#FCD535)와 그 상태값(`primary-active`, `primary-disabled`)
- 타이포그래피 폰트 패밀리(`--font-nova`, `--font-plex`)와 크기/굵기 위계
- 트레이딩 의미 색(`trading-up`/`trading-down`)과 포커스 링(`info`)
- 반경 스케일(`--radius-*`)과 버튼 위에 얹히는 텍스트 색(`--on-primary`는 항상 어두운 잉크 —
  design.md의 "노랑 위 검정은 이 시스템의 시그니처" 규칙을 두 테마 모두에서 지킨다)

**테마마다 바뀌는 것**: 캔버스, 카드 표면, 보조 표면(입력창/보조 버튼 배경), 헤어라인, 강한 텍스트,
본문 텍스트, 진하게 강조된 muted 텍스트.

## 구현 방식

`static/style.css`는 `:root`에 다크 값을 기본으로 두고, `:root[data-theme="light"]`에서 아래
토큰만 재정의한다. `<html>`의 `data-theme` 속성은 헤더의 테마 토글 버튼(`#themeToggle`,
`static/app.js`)이 클릭 시 전환하며, 선택값은 `localStorage("splitexcelweb-theme")`에 저장해
새로고침 후에도 유지된다. 깜빡임(FOUC)을 막기 위해 `index.html` `<head>`의 인라인 스크립트가
CSS 로드 전에 저장된 값을 미리 적용한다.

## 색상 토큰

| 토큰 | 다크(기본) | 라이트 | 역할 |
|---|---|---|---|
| `--canvas` | `#0b0e11` | `#ffffff` | 페이지 바닥면 (design.md의 `canvas-dark` / `canvas-light`) |
| `--surface-card` | `#1e2329` | `#ffffff` | 카드 배경. 라이트에서는 캔버스와 동일한 흰색을 쓰고 `--hairline` 테두리로만 구분한다 — design.md의 라이트 거래성 컴포넌트(`steps-card`, `price-chart-card` 등)가 실제로 이렇게 되어 있다(색 블록이 아니라 헤어라인으로 구분) |
| `--surface-elevated` | `#2b3139` | `#f5f5f5` | 입력창·보조 버튼 배경 (design.md `surface-strong-light`) |
| `--hairline` | `#2b3139` | `#eaecef` | 1px 테두리/구분선 (design.md `hairline-on-light`) |
| `--border-strong` | `#3a4048` | `#cdd1d6` | 보조 버튼 hover, 강조 테두리 (design.md `border-strong`) |
| `--text-strong` | `#ffffff` | `#181a20` | 헤딩 등 최고 대비 텍스트 (design.md `on-dark` / `ink`) |
| `--body-text` | `#eaecef` | `#181a20` | 기본 본문 텍스트 (라이트는 `ink` 재사용 — design.md: "라이트 모드 본문 텍스트는 ink 토큰을 재사용") |
| `--muted` | `#707a8a` | `#707a8a` | 캡션/보조 라벨. 두 테마 모두 동일(design.md: "라이트와 다크 캔버스 모두에서 동작") |
| `--muted-strong` | `#929aa5` | `#55606e` | 2단계 muted. 라이트에서는 흰 배경 위 가독성을 위해 더 어둡게 조정(design.md 값 그대로 쓰면 흰 배경에서 대비가 약해서 splitExcelWeb에서 자체 보정) |

브랜드/의미 토큰(테마 공통, 값 불변):

| 토큰 | 값 | 역할 |
|---|---|---|
| `--primary` | `#fcd535` | 브랜드 액센트, 주요 CTA |
| `--primary-active` | `#f0b90b` | 주요 CTA 눌림 상태 |
| `--primary-disabled` | `#3a3a1f` | 비활성 CTA (다크 캔버스 기준 — 라이트에서 쓸 일이 생기면 재검토 필요, 알려진 공백 참고) |
| `--ink` / `--on-primary` | `#181a20` | 노란 배경 위 텍스트. 항상 어두운 잉크 |
| `--trading-up` | `#0ecb81` | 상태/성공 계열에 재활용하지 말 것(design.md 규칙과 동일) |
| `--trading-down` | `#f6465d` | 위 항목과 동일 |
| `--info` | `#3b82f6` | 포커스 링 |

## Do's and Don'ts

### Do
- 새 컴포넌트를 추가할 때 다크 전용 색을 직접 쓰지 말고 항상 위 CSS 변수(`var(--canvas)` 등)를
  통해 참조할 것 — 그래야 라이트 테마에서도 자동으로 맞는 색이 적용된다.
- 카드/모달처럼 표면이 필요한 요소는 라이트에서 `--surface-card`(흰색)와 `--hairline`(테두리)
  조합으로 구분한다는 원칙을 유지할 것. 라이트 모드에 그림자를 새로 추가하지 말 것(design.md의
  "평면 표면" 철학 유지).
- 반투명 오버레이(포커스 링, hover 틴트, 선택 셀 하이라이트 등 `rgba()`)는 테마와 무관하게 대체로
  잘 동작하므로 대부분 그대로 재사용 가능 — 새로 추가할 때도 알파 블렌딩을 우선 고려할 것.

### Don't
- `--primary`(노란색)를 라이트 테마라고 다른 색으로 바꾸지 말 것. 브랜드 액센트는 테마와 무관하게
  고정이다.
- `--body-text`/`--text-strong`을 라이트에서 회색 계열로 낮추지 말 것 — design.md 규칙대로 `ink`
  (`#181a20`)를 그대로 재사용해 다크의 흰색 텍스트만큼 확실한 대비를 유지해야 한다.
- 라이트 전용 컴포넌트를 만들 때 다크에서 쓰는 진한 카드 배경(`#1e2329` 등)을 하드코딩하지 말 것.

## 알려진 공백

- `--primary-disabled`(#3a3a1f)는 다크 캔버스를 전제로 만들어진 값이라 라이트 배경 위 비활성
  버튼에는 대비가 약할 수 있다. 현재 UI에 비활성 상태의 주요(노란) 버튼이 없어 아직 실사용
  검증은 못 했다 — 필요해지면 라이트 전용 값을 새로 정의할 것.
- 그리드 미리보기(범위 선택 뷰어)의 스크롤바, 네이티브 `<select>` 드롭다운 등 브라우저 기본
  렌더링 요소는 `color-scheme`(다크는 `dark`, 라이트는 `light`)에 맡기고 있어 브라우저/OS별로
  스타일이 조금씩 다를 수 있다.
