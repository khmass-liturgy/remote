# vendor/ieum-pptx.js 빌드 방법

`vendor/ieum-pptx.js`는 [이음오피스(IeumOffice)](https://github.com/pblsketch/ieumoffice)의
파워포인트 엔진(`packages/pptx-engine`, `packages/pptx-render`, `apps/slides`의 Konva 렌더러)을
브라우저에서 읽기 전용 뷰어로 쓰도록 esbuild로 묶은 파일입니다. (Apache-2.0, 고지문은 `vendor/ieumoffice/`)

기준 커밋: `58031f1a58bc52aa797f426bb8ef3c07d1929be2`

```bash
git clone https://github.com/pblsketch/ieumoffice.git
mkdir build && cd build && cp ../remote/tools/ieum-pptx/* .
npm init -y
npm i --ignore-scripts esbuild react@19 react-dom@19 konva@9 react-konva@19 opentype.js@2 bidi-js jszip fast-xml-parser@5 polygon-clipping buffer
mkdir -p node_modules/@genoffice
for p in ../ieumoffice/packages/*; do ln -sfn "$(realpath $p)" node_modules/@genoffice/$(basename $p); done
IEUM_SRC=$(realpath ../ieumoffice) node build.mjs      # → ieum-pptx.js
```

`entry.tsx`가 이 저장소에서 쓴 연결 코드이고(파일 열기, 슬라이드 그리기, 동영상 위치·노트 추출),
나머지는 이음오피스 원본 그대로입니다.
