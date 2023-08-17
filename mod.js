const rcedit = require('rcedit')
var config = require('./package.json');

rcedit('C:\\Users\\Hello\\.pkg-cache\\v3.4\\built-v18.5.0-win-x64' , {
  icon: "./my.ico",
  "version-string" :{
    ProductName: config.name,
    LegalCopyright: config.copyright,
    FileDescription: config.description,
    OriginalFilename: ""
  },
  "file-version": config.version,
  "product-version": config.version,
},(err) => {
  console.log(err)
})