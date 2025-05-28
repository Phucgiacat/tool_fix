document.getElementById('excelFile').addEventListener('change', function (e) {
  const file = e.target.files[0];
  const formData = new FormData();
  formData.append("file", file);

  fetch("http://127.0.0.1:5000/upload", {
    method: "POST",
    body: formData
  })
    .then(res => res.json())
    .then(data => {
      const reader = new FileReader();
      reader.onload = function (event) {
        const cxCharColors = {};
        const cxCharTexts = {};  // để lưu lại từng chữ ở C sau khi gọi API
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array', cellStyles: true });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const range = XLSX.utils.decode_range(worksheet['!ref']);

        const cxFetchTasks = [];

        // Giai đoạn 1: gọi API để lấy từng chữ và màu của cột C
        for (let row = range.s.r; row <= range.e.r; row++) {
          const rowNumber = row + 1;
          const cellAddress = `C${rowNumber}`;
          const cell = worksheet[cellAddress];
          if (cell && rowNumber >= 2) {
            const task = fetch(`http://127.0.0.1:5000/sequence?char=${cellAddress}`)
              .then(res => res.json())
              .then(data => {
                const colors = [];
                const chars = [];
                data.forEach(item => {
                  const itemColor = item.Color && item.Color !== 'default'
                    ? `#${item.Color.replace(/^FF/, '')}`
                    : '#000000';
                  colors.push(itemColor);
                  chars.push(item.Text);
                });
                cxCharColors[rowNumber] = colors;
                cxCharTexts[rowNumber] = chars;
              })
              .catch(err => {
                console.error(`Lỗi API cho ô ${cellAddress}:`, err);
                cxCharColors[rowNumber] = [];
                cxCharTexts[rowNumber] = [];
              });
            cxFetchTasks.push(task);
          }
        }

        Promise.all(cxFetchTasks).then(() => {
          let tableMatrix = [];

          for (let row = range.s.r; row <= range.e.r; row++) {
            let rowCells = [];

            for (let col = range.s.c; col <= range.e.c; col++) {
              const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
              const cell = worksheet[cellAddress];
              const columnLetter = cellAddress.match(/^[A-Z]+/)[0];
              const rowNumber = row + 1;

              if (cell) {
                const cellValue = cell.v ?? "";

                if (columnLetter === 'C' && rowNumber >= 2) {
                  const chars = cxCharTexts[rowNumber] || [];
                  const colors = cxCharColors[rowNumber] || [];
                  const html = chars.map((ch, i) =>
                    `<span class="char-cx" data-row="${rowNumber}" data-index="${i}" style="color:${colors[i]};">${ch}</span>`
                  ).join('');
                  rowCells.push(Promise.resolve(`<td>${html}</td>`));
                }

                else if (columnLetter === 'E' && rowNumber >= 2) {
                  const words = String(cellValue).split(' ');
                  const colors = cxCharColors[rowNumber] || [];
                  const html = words.map((word, i) =>
                    `<span class="char-ex" data-row="${rowNumber}" data-index="${i}" data-char="${word}" style="color:${colors[i] || '#000000'}; cursor:pointer;">${word}</span>`
                  ).join(' ');
                  rowCells.push(Promise.resolve(`<td>${html}</td>`));
                }

                else {
                  let color = "000000";
                  if (cell.s?.font?.color?.rgb) {
                    color = cell.s.font.color.rgb;
                  }
                  rowCells.push(Promise.resolve(`<td style="color:#${color}">${cellValue}</td>`));
                }
              } else {
                rowCells.push(Promise.resolve('<td></td>'));
              }
            }

            tableMatrix.push(Promise.all(rowCells));
          }

          Promise.all(tableMatrix).then(rows => {
            let html = '<table border="1" cellspacing="0" cellpadding="5">';
            rows.forEach(cells => {
              html += '<tr>' + cells.join('') + '</tr>';
            });
            html += '</table>';
            document.getElementById("output").innerHTML = html;

            // Gắn click cho chữ ở cột E để sửa chữ ở cột C
            document.querySelectorAll('.char-ex').forEach(span => {
              span.addEventListener('click', () => {
                const originalText = span.dataset.char;
                const row = span.dataset.row;
                const index = span.dataset.index;

                fetch(`http://127.0.0.1:5000/suggest?char=${encodeURIComponent(originalText)}`)
                  .then(r => r.json())
                  .then(data => {
                    const suggestions = data.suggestions;
                    if (!suggestions.length) return alert(`Không có gợi ý cho '${originalText}'`);

                    const select = document.createElement('select');
                    suggestions.forEach(s => {
                      const option = document.createElement('option');
                      option.value = s;
                      option.textContent = s;
                      select.appendChild(option);
                    });

                    select.addEventListener('change', () => {
                      const newChar = select.value;
                      const cxSpan = document.querySelector(`.char-cx[data-row="${row}"][data-index="${index}"]`);
                      if (cxSpan) {
                        cxSpan.textContent = newChar;
                        cxSpan.style.color = 'green';
                        cxSpan.style.fontWeight = 'bold';
                      }
                      select.remove();
                    });

                    span.after(select);
                  })
                  .catch(err => {
                    console.error('❌ Lỗi gợi ý:', err);
                    alert(`Không thể lấy gợi ý cho '${originalText}'`);
                  });
              });
            });
          });
        });
      };
      reader.readAsArrayBuffer(file);
    })
    .catch(err => {
      console.error("Lỗi:", err);
    });
});