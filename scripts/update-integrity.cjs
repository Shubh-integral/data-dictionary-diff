const fs = require('node:fs');
const path = require('node:path');
const {createHash} = require('node:crypto');
const root = path.resolve(__dirname, '..');
const file = path.join(root, 'index.html');
let html = fs.readFileSync(file, 'utf8');
html = html.replace(/<(script|link)\b[^>]*>/g, tag => {
  const resource = tag.match(/(?:src|href)="([^"]+)"/);
  if (!resource || !/\.(js|css)$/.test(resource[1])) return tag;
  const digest = createHash('sha384').update(fs.readFileSync(path.join(root, resource[1]))).digest('base64');
  return tag.replace(/ integrity="[^"]*"/g, '').replace(/>$/, ` integrity="sha384-${digest}">`);
});
fs.writeFileSync(file, html);
