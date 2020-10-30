import * as ExcelJS from 'exceljs';

async function main(){
  const template_filename = 'template/rirekisho.xlsx';
  const output_file = 'output.xlsx';

  const photo_filename = 'template/photo.png';
  const name_range = "D7";
  const photo_range = 'H3:H10';

  // エクセルのテンプレートを開く
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(template_filename);

  // 最初のシートを取り出す
  const worksheet = workbook.worksheets[0];

  // 文字を書く
  worksheet.getCell(name_range).value = 'テスト　テスト';

  // 画像を貼る
  const imageId1 = workbook.addImage({
    filename: photo_filename,
    extension: 'png',
  });
  worksheet.addImage(imageId1, photo_range);

  // エクセルを保存
  await workbook.xlsx.writeFile(output_file);
}

main();
