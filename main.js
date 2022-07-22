var XLSX = require("xlsx")
var fs = require('fs')

var workbook = XLSX.readFile("arquivo.xlsx")
var sheet_name_list = workbook.SheetNames

var headers = {}
var data = []

sheet_name_list.forEach(function (y) {
  var worksheet = workbook.Sheets[y]
  for (z in worksheet) {
    if (z[0] === "!") continue
    var col = z.substring(0, 1)
    var row = parseInt(z.substring(1))
    var value = worksheet[z].v
    if (row == 1) {
      headers[col] = value
      continue
    }
    if (!data[row]) data[row] = {}
    data[row][headers[col]] = value
  }
  data.shift()
  data.shift()
})

var logStream = fs.createWriteStream('Dados.txt', { flags: 'a' });
for (position in data) {
  var values = []
  if (data[position].Jan && data[position].Jan != 0.0) {
    values.push(data[position].Jan)
  }
  if (data[position].Fev && data[position].Fev != 0.0) {
    values.push(data[position].Fev)
  }
  if (data[position].Mar && data[position].Mar != 0.0) {
    values.push(data[position].Mar)
  }
  if (data[position].Abr && data[position].Abr != 0.0) {
    values.push(data[position].Abr)
  }
  if (data[position].Mai && data[position].Mai != 0.0) {
    values.push(data[position].Mai)
  }
  if (data[position].Jun && data[position].Jun != 0.0) {
    values.push(data[position].Jun)
  }
  if (data[position].Jul && data[position].Jul != 0.0) {
    values.push(data[position].Jul)
  }
  if (data[position].Ago && data[position].Ago != 0.0) {
    values.push(data[position].Ago)
  }
  if (data[position].Set && data[position].Set != 0.0) {
    values.push(data[position].Set)
  }
  if (data[position].Out && data[position].Out != 0.0) {
    values.push(data[position].Out)
  }
  if (data[position].Nov && data[position].Nov != 0.0) {
    values.push(data[position].Nov)
  }
  if (data[position].Dez && data[position].Dez != 0.0) {
    values.push(data[position].Dez)
  }

  var uniq = values
    .map((value) => {
      return {
        count: 1,
        value: value
      }
    })
    .reduce((a, b) => {
      a[b.value] = (a[b.value] || 0) + b.count
      return a
    }, {})
  var duplicates = Object.keys(uniq).filter((a) => uniq[a] > 1)
  if (duplicates.length != 0) {
    console.log(`${data[position].Mun}, ${data[position].Reg}, ${data[position].Ano}, ${duplicates}`)
    logStream.write(`${data[position].Mun}, ${data[position].Reg}, ${data[position].Ano}, ${duplicates}\n`);
  }
}