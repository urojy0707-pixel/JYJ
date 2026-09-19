const React = require('react');
const ReactDOMServer = require('react-dom/server');
const sharp = require('sharp');

async function iconData(Comp, color, px) {
  px = px || 320;
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(Comp, { color: '#' + color, size: px })
  );
  const buf = await sharp(Buffer.from(svg))
    .resize(px, px, { fit: 'contain', background: { r: 0, g: 0, b: 0, alpha: 0 } })
    .png()
    .toBuffer();
  return 'image/png;base64,' + buf.toString('base64');
}
module.exports = { iconData };
