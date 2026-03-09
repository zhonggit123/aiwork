const fs = require('fs');
const path = require('path');

const svgPath = path.join(__dirname, 'icon.svg');
const svg = fs.readFileSync(svgPath, 'utf8');

let Resvg;
try {
  Resvg = require('@resvg/resvg-js').Resvg;
} catch (e) {
  console.error('请先安装: npm install @resvg/resvg-js');
  process.exit(1);
}

const sizes = [16, 32, 48, 128];
for (const size of sizes) {
  const resvg = new Resvg(svg, { fitTo: { mode: 'width', value: size } });
  const png = resvg.render().asPng();
  fs.writeFileSync(path.join(__dirname, `icon${size}.png`), png);
  console.log('生成 icon' + size + '.png');
}
console.log('完成');
