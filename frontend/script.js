document.getElementById('excelFile').addEventListener('change', function(e) {
  const file = e.target.files[0];
  const reader = new FileReader();

  reader.onload = function(event) {
    const data = new Uint8Array(event.target.result);
    const workbook = XLSX.read(data, { type: 'array', cellStyles: true });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const range = XLSX.utils.decode_range(worksheet['!ref']);

    let html = '<table border="1" cellspacing="0" cellpadding="5">';
    for (let row = range.s.r; row <= range.e.r; row++) {
      html += '<tr>';
      for (let col = range.s.c; col <= range.e.c; col++) {
        const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
        const cell = worksheet[cellAddress];

        if (cell) {
          let text = cell.v ?? "";
          let color = "000000";

          if (cell.s?.font?.color?.rgb) {
            color = cell.s.font.color.rgb;
            console.log(`✅ Có màu ở ô ${cellAddress}: ${text} → #${color}`);
          }

          html += `<td style="color:#${color}">${text}</td>`;
        } else {
          html += '<td></td>';
        }
      }
      html += '</tr>';
    }
    html += '</table>';

    document.getElementById("output").innerHTML = html;
  };

  reader.readAsArrayBuffer(file);
});
