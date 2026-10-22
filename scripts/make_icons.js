// Renders the slide icons to white PNGs in assets/icons.
// Needs: npm install react-icons react react-dom sharp
const path = require("path");
const fs = require("fs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa6");

const ICONS = {
  display: fa.FaRectangleAd,
  search: fa.FaMagnifyingGlassDollar,
  video: fa.FaCirclePlay,
  sms: fa.FaCommentSms,
  inapp: fa.FaMobileScreenButton,
  qr: fa.FaQrcode,
  geo: fa.FaLocationDot,
  seo: fa.FaMagnifyingGlassChart,
};

const outDir = path.join(__dirname, "..", "assets", "icons");
fs.mkdirSync(outDir, { recursive: true });

(async () => {
  for (const [name, Icon] of Object.entries(ICONS)) {
    const svg = ReactDOMServer.renderToStaticMarkup(
      React.createElement(Icon, { color: "#FFFFFF", size: 512 })
    );
    await sharp(Buffer.from(svg)).resize(512, 512).png().toFile(path.join(outDir, `${name}.png`));
    console.log(`wrote ${name}.png`);
  }
})();
