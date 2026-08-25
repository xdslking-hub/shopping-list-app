const React = require('react');
const ReactDOMServer = require('react-dom/server');
const sharp = require('sharp');
const fa = require('react-icons/fa');
const pptxgen = require('pptxgenjs');

const BG = '0B1220';       // deep navy-black (dominant)
const CARD = '17223B';     // slate navy (secondary, cards)
const ACCENT = '00E5FF';   // electric cyan (sharp accent)
const TEXT_MUTED = '9AA7C7';
const TEXT_BODY = 'D7DEF0';

async function iconToPngBase64(IconComponent, colorHex, px = 256) {
  const svgMarkup = ReactDOMServer.renderToStaticMarkup(
    React.createElement(IconComponent, { color: `#${colorHex}`, size: px })
  );
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${px}" height="${px}" viewBox="0 0 512 512">${svgMarkup.replace(/<svg[^>]*>|<\/svg>/g, '')}</svg>`;
  const buf = await sharp(Buffer.from(svg)).resize(px, px).png().toBuffer();
  return 'image/png;base64,' + buf.toString('base64');
}

async function main() {
  const pres = new pptxgen();
  pres.layout = 'LAYOUT_WIDE'; // 13.33 x 7.5

  const slide = pres.addSlide();
  slide.background = { color: BG };

  // Title
  slide.addText('AI 기술 트렌드', {
    x: 0.6, y: 0.45, w: 9, h: 0.8,
    fontFace: 'Cambria', fontSize: 40, bold: true, color: 'FFFFFF',
    margin: 0,
  });
  slide.addText('2026년 하반기, 산업 전반에서 반복적으로 관찰되는 6가지 방향', {
    x: 0.6, y: 1.18, w: 10.5, h: 0.4,
    fontFace: 'Calibri', fontSize: 14, italic: true, color: TEXT_MUTED,
    margin: 0,
  });

  const trends = [
    {
      icon: fa.FaRobot,
      title: '에이전틱 AI',
      desc: '단일 응답형 챗봇을 넘어, 여러 단계를 스스로 계획·실행하는 자율 에이전트로 무게중심 이동',
    },
    {
      icon: fa.FaLayerGroup,
      title: '멀티모달 통합',
      desc: '텍스트·이미지·음성·영상을 하나의 모델이 함께 이해·생성하는 방향으로 표준화',
    },
    {
      icon: fa.FaMicrochip,
      title: '경량화·온디바이스',
      desc: '소형·경량 모델이 발전하며 클라우드 의존을 줄이는 엣지·로컬 실행 사례 확대',
    },
    {
      icon: fa.FaBalanceScale,
      title: '규제·거버넌스 강화',
      desc: 'EU AI Act 등 지역별 규제가 시행 단계에 진입하며 컴플라이언스 대응이 필수 과제로 부상',
    },
    {
      icon: fa.FaBolt,
      title: '추론 비용 절감',
      desc: '모델 압축·효율화 기술 발전으로 동일 성능 대비 추론(inference) 비용이 지속적으로 하락',
    },
    {
      icon: fa.FaBriefcase,
      title: '기업 도입 가속화',
      desc: 'RAG·코파일럿·워크플로우 자동화를 중심으로 실무 적용 사례가 실험 단계를 지나 확산 단계로',
    },
  ];

  const marginX = 0.6;
  const gap = 0.35;
  const cols = 3, rows = 2;
  const gridTop = 1.9;
  const gridBottom = 7.05;
  const cardW = (13.33 - marginX * 2 - gap * (cols - 1)) / cols;
  const cardH = (gridBottom - gridTop - gap * (rows - 1)) / rows;

  const circleSize = 0.55;

  for (let i = 0; i < trends.length; i++) {
    const col = i % cols;
    const row = Math.floor(i / cols);
    const x = marginX + col * (cardW + gap);
    const y = gridTop + row * (cardH + gap);

    // card background
    slide.addShape('roundRect', {
      x, y, w: cardW, h: cardH,
      rectRadius: 0.08,
      fill: { color: CARD },
      line: { type: 'none' },
      shadow: {
        type: 'outer', color: '000000', opacity: 0.35,
        blur: 6, offset: 3, angle: 90,
      },
    });

    // icon circle
    const circleX = x + 0.3;
    const circleY = y + 0.3;
    slide.addShape('ellipse', {
      x: circleX, y: circleY, w: circleSize, h: circleSize,
      fill: { color: ACCENT },
      line: { type: 'none' },
    });

    const iconPng = await iconToPngBase64(trends[i].icon, BG, 256);
    const iconPad = 0.14;
    slide.addImage({
      data: iconPng,
      x: circleX + iconPad, y: circleY + iconPad,
      w: circleSize - iconPad * 2, h: circleSize - iconPad * 2,
    });

    // title
    slide.addText(trends[i].title, {
      x: circleX + circleSize + 0.15, y: circleY - 0.05, w: cardW - (circleSize + 0.15 + 0.25), h: 0.6,
      fontFace: 'Calibri', fontSize: 16, bold: true, color: 'FFFFFF',
      valign: 'middle', margin: 0,
    });

    // description
    slide.addText(trends[i].desc, {
      x: x + 0.3, y: circleY + circleSize + 0.16, w: cardW - 0.6, h: cardH - circleSize - 0.6,
      fontFace: 'Calibri', fontSize: 12.5, color: TEXT_BODY,
      valign: 'top', margin: 0, lineSpacingMultiple: 1.15,
    });
  }

  await pres.writeFile({ fileName: 'AI_기술_트렌드_요약.pptx' });
  console.log('done');
}

main().catch((e) => { console.error(e); process.exit(1); });
