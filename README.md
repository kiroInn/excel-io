<div align="center">
<h1>Excel IO</h1>

<br />
<a href="http://kiroinn.github.io/excelio">
  <img
    height="80"
    width="80"
    alt="lizard"
    src="https://raw.githubusercontent.com/kiroInn/excel-io/master/src/assets/logo.png"
  />
</a>
<br />
<br />
<p>excel io is excel productivity tool, support workbook split and collect any cell evaluation.</p>
<br />

[**Online Preview**][excelio]

</div>
<hr />

## Tech Stack
- vue-cli
- vue
- vue-resource
- FileSaver
- exceljs
- jszip
- testing-library
- ui design: https://www.manypixels.co
  
## Project setup
```
npm install
```

### Compiles and hot-reloads for development
```
npm run serve
```

### Compiles and minifies for production
```
npm run build
```

### Lints and fixes files
```
npm run lint
```

## Enhancement exceljs copy sheet
sed -i '' 's/c=e.drawing;s=t.media\[o.imageId\]/c=e.drawing;s=t.media\[o.imageId\]||{}/g' node_modules/exceljs/dist/exceljs.min.js
sed -i '' 's/\,u\[c.rels.length\]=f\,/\,u\[c.rels.length\]=f\,s\&\&/g' node_modules/exceljs/dist/exceljs.min.js

## License

[MIT][license]

<!-- prettier-ignore-start -->
[excelio]: http://kiroinn.github.io/excelio
[license]: https://github.com/kiroInn/excel-io/blob/master/LICENSE.MD
<!-- prettier-ignore-end -->