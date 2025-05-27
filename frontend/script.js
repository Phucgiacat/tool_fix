// Hàm ARGB -> CSS
function argbToHex(argb) {
  if (!argb || argb === 'default') return '#000000';
  return '#' + argb.slice(2);
}

// Giải mã unicode escape: \\uXXXX hoặc \\UXXXXXXXX -> ký tự thật
function decodeUnicodeEscape(str) {
  return str.replace(/\\u([\da-f]{4})/gi, (_, grp) =>
    String.fromCharCode(parseInt(grp, 16))
  ).replace(/\\U([\da-f]{8})/gi, (_, grp) =>
    String.fromCodePoint(parseInt(grp, 16))
  );
}

function renderSentences(data) {
  const container = document.getElementById('sentence-container');
  container.innerHTML = '';

  const sentences = data["0"];
  sentences.forEach((sentence, sIdx) => {
    const sentenceDiv = document.createElement('div');
    sentenceDiv.className = 'sentence';

    sentence.forEach((word, wIdx) => {
      const span = document.createElement('span');
      span.className = 'word';

      const realText = decodeUnicodeEscape(word.Text || '');
      span.textContent = realText;
      span.style.color = argbToHex(word.Color);
      if (word.Font && word.Font !== 'default') {
        span.style.fontFamily = word.Font;
      }

      // Khi click -> gọi API /suggest?char=...
      span.onclick = () => {
        fetch(`http://127.0.0.1:5000/suggest?char=${encodeURIComponent(realText)}`)
          .then(res => res.json())
          .then(data => {
            const similars = data.suggestions;
            if (!similars || similars.length === 0) {
              alert(`Không có gợi ý cho "${realText}"`);
              return;
            }

            const select = document.createElement('select');
            select.className = 'dropdown';
            similars.forEach(sim => {
              const option = document.createElement('option');
              option.value = sim;
              option.textContent = sim;
              select.appendChild(option);
            });

            select.onchange = () => {
              span.textContent = select.value;
              span.appendChild(select);
            };

            span.appendChild(select);
          })
          .catch(err => {
            alert('Lỗi khi gọi API: ' + err.message);
          });
      };

      sentenceDiv.appendChild(span);
    });

    container.appendChild(sentenceDiv);
  });
}

document.getElementById('excelFile').addEventListener('change', function (e) {
  const file = e.target.files[0];
  const reader = new FileReader();

  reader.onload = function (event) {
    const data = new Uint8Array(event.target.result);
    const workbook = XLSX.read(data, { type: 'array', cellStyles: true });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const range = XLSX.utils.decode_range(worksheet['!ref']);

    // Tạo ma trận lưu từng ô tạm thời
    let tableMatrix = [];

    for (let row = range.s.r; row <= range.e.r; row++) {
      let rowCells = [];

      for (let col = range.s.c; col <= range.e.c; col++) {
        const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
        const cell = worksheet[cellAddress];
        if (cell) {
          const cellValue = cell.v ?? "";
          const columnLetter = cellAddress.match(/^[A-Z]+/)[0];
          const rowNumber = parseInt(cellAddress.match(/\d+$/)[0]);

          if (columnLetter === 'D' && rowNumber >= 2) {
            // Gọi API đồng bộ bằng cách chờ Promise
            rowCells.push(
              fetch(`http://127.0.0.1:5000/sequence?char=${cellAddress}`)
                .then(response => response.json())
                .then(data => {
                  let cellHtml = '';
                  data.forEach(item => {
                    console.log(item.Text)
                    const font = item.Font ? Array.from(item.Font).join(', ') : 'default';
                    const color = item.Color && item.Color !== 'default'
                      ? `#${item.Color.replace(/^FF/, '')}`
                      : '#000000';
                      cellHtml += `<span class="char-clickable" 
                                        data-char="${item.Text}" 
                                        style="font-family:${font}; color:${color}; cursor:pointer;">
                                      ${item.Text}
                                  </span>`;

                  });
                  return `<td>${cellHtml}</td>`;
                })


                .catch(err => {
                  console.error(`❌ Lỗi API cho ô ${cellAddress}:`, err);
                  return `<td style="color:#000000">${cellValue}</td>`;
                })
            );
          } else {
            // Ô thường
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

    // Kết hợp tất cả hàng và render sau khi hoàn tất
    Promise.all(tableMatrix).then(rows => {
      let html = '<table border="1" cellspacing="0" cellpadding="5">';
      rows.forEach(cells => {
        html += '<tr>' + cells.join('') + '</tr>';
      });
      html += '</table>';
      document.getElementById("output").innerHTML = html;
      // Gắn sự kiện click vào từng chữ có class 'char-clickable'
      document.querySelectorAll('.char-clickable').forEach(span => {
        span.addEventListener('click', () => {
          const originalText = span.dataset.char;

          fetch(`http://127.0.0.1:5000/suggest?char=${encodeURIComponent(originalText)}`)
            .then(r => r.json())
            .then(data => {
              const suggestions = data.suggestions;
              if (!suggestions.length) {
                alert(`Không có gợi ý cho '${originalText}'`);
                return;
              }

              // Tạo dropdown chọn gợi ý
              const select = document.createElement('select');
              select.style.marginLeft = '4px';
              select.style.fontSize = 'inherit';
              select.style.fontFamily = 'inherit';

              // Thêm option vào dropdown
              suggestions.forEach(s => {
                const option = document.createElement('option');
                option.value = s;
                option.textContent = s;
                select.appendChild(option);
              });

              // Khi chọn → thay thế chữ gốc
              select.addEventListener('change', () => {
                span.textContent = select.value;
                span.dataset.char = select.value;
                span.style.color = 'green';
                span.style.fontWeight = 'bold';
                select.remove();
              });


              // Chèn dropdown bên cạnh span
              span.after(select);
            })
            .catch(err => {
              console.error('❌ Lỗi gợi ý:', err);
              alert(`Không thể lấy gợi ý cho '${originalText}'`);
            });
        });
      });




    });
  };

  reader.readAsArrayBuffer(file);
  
});

