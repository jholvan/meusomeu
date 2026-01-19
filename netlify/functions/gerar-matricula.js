const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");

exports.handler = async (event) => {
  const dados = JSON.parse(event.body);

  const workbook = new ExcelJS.Workbook();
  const filePath = path.join(__dirname, "../../PRE.xlsx");

  await workbook.xlsx.readFile(filePath);
  const sheet = workbook.getWorksheet(1);

  sheet.getCell("B7").value = dados.curso;
  sheet.getCell("B8").value = dados.turno;
  sheet.getCell("D7").value = dados.dia;

  sheet.getCell("B10").value = dados.nome;
  sheet.getCell("B11").value = dados.cpf;
  sheet.getCell("B19").value = dados.numero;

  sheet.getCell("B32").value = dados.parcelas;
  sheet.getCell("E30").value = dados.desconto;
  sheet.getCell("E32").value = dados.valorParcela;

  const buffer = await workbook.xlsx.writeBuffer();

  return {
    statusCode: 200,
    headers: {
      "Content-Type":
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "Content-Disposition": "attachment; filename=PRE.xlsx"
    },
    body: buffer.toString("base64"),
    isBase64Encoded: true
  };
};