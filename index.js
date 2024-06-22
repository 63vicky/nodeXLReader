var http = require("http");
var xlsx = require("xlsx");
var formidable = require("formidable");
const { readFile, existsSync, mkdirSync, rename } = require("fs");
const { mkdir } = require("fs/promises");
http
  .createServer(function (request, response) {
    if (request.url == "/excelRead") {
      uploadXL(request, response).then(() => {
        let workBook = xlsx.readFile("SaleData.xlsx");
        let sheetNames = workBook.SheetNames;
        let xlData = xlsx.utils.sheet_to_html(workBook.Sheets[sheetNames[0]]);
        response.end(xlData);
      });
    } else {
      response.writeHead(200, { "Content-Type": "text/html" });
      readFile("index.html", (err, data) => {
        if (err) throw err;
        response.write(data);
        response.end("");
      });
    }
  })
  .listen(8083);

const uploadXL = async (request, response) => {
  var form = new formidable.IncomingForm();
  await form.parse(request, (err, fields, files) => {
    files.XLFile.forEach((file) => {
      var oldpath = file.filepath;
      var newpath = "C:/Vivek/Node/nodeXLReader/nodeFileUpload/";
      if (!existsSync(newpath)) {
        mkdir(newpath).then(() => {
          rename(oldpath, newpath + file.originalFilename, function (err) {
            if (err) err;
            response.write("File uploaded and moved!");
            response.end();
          });
        });
      } else {
        rename(oldpath, newpath + file.originalFilename, function (err) {
          if (err) err;
          response.write("File uploaded and moved!");
          response.end();
        });
      }
    });
  });
};

console.log("Server running at http://127.0.0.1:8083/");
