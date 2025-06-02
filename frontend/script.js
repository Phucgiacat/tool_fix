document.getElementById('excelFile').addEventListener('change', function (e) {
  const file = e.target.files[0];
  const inputValue = document.getElementById("path_image").value;
  const formData = new FormData();
  formData.append("file", file);
  formData.append("path_folder", inputValue);

  fetch("http://127.0.0.1:5000/upload", {
    method: "POST",
    body: formData
  })
    .then(res => res.json())
    .then(data => {
      const reader = new FileReader();
      reader.onload = function (event) {
        const cxCharColors = {};
        const cxCharTexts = {}; 
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
                rowCells.push(Promise.resolve(`<td style="color:#000000">${cellValue}</td>`));
                }
              } else {
                rowCells.push(Promise.resolve('<td></td>'));
              }
            }

            tableMatrix.push(Promise.all(rowCells));
          }

          Promise.all(tableMatrix).then(rows => {
            // Hiển thị ảnh khi hover vào dòng (x >= 2)
            const aColumnCache = {};
            for (let row = range.s.r + 1; row <= range.e.r; row++) {
              const rowNumber = row + 1;
              const aCell = worksheet[`A${rowNumber}`];
              aColumnCache[row] = aCell?.v ?? '';
            }

            // Gắn lại bảng kèm data-img
            let html = '<table border="1" cellspacing="0" cellpadding="5">';
            rows.forEach((cells, i) => {
              const dataImg = i >= 1 ? ` data-img="${aColumnCache[i] || ''}"` : '';
              html += `<tr${dataImg}>${cells.join('')}</tr>`;
            });
            html += '</table>';
            document.getElementById("output").innerHTML = html;
            // Thêm container cố định bên phải nếu chưa có
            if (!document.getElementById('side-image-box')) {
              const sideBox = document.createElement('div');
              sideBox.id = 'side-image-box';
              sideBox.style.position = 'fixed';
              sideBox.style.top = '10px';
              sideBox.style.right = '10px';
              sideBox.style.width = '220px';
              sideBox.style.padding = '6px';
              sideBox.style.border = '1px solid #ccc';
              sideBox.style.background = '#fff';
              sideBox.style.boxShadow = '0 0 8px rgba(0,0,0,0.2)';
              sideBox.style.display = 'none';
              sideBox.style.zIndex = 9999;

              const title = document.createElement('div');
              sideBox.style.display = 'none';
              sideBox.style.zIndex = 9999;
              sideBox.style.display = 'flex';            // thêm dòng này
              sideBox.style.flexDirection = 'column';    // để dồn dọc
              sideBox.style.alignItems = 'center';       // canh giữa ngang
              sideBox.style.justifyContent = 'center';   // canh giữa dọc
              title.style.marginBottom = '6px';
              title.style.fontWeight = 'bold';
              title.textContent = 'Ảnh dòng đang trỏ';

              const img = document.createElement('img');
              img.id = 'side-image';
              img.style.maxWidth = '100%';
              img.style.maxHeight = '300px';
              img.style.border = '1px solid #eee';

              sideBox.appendChild(title);
              sideBox.appendChild(img);
              document.body.appendChild(sideBox);
            }

            // Gắn sự kiện hover
              // Gắn sự kiện hover + tô nền hàng
              document.querySelectorAll('tr[data-img]').forEach(tr => {
                tr.addEventListener('mouseenter', () => {
                  const id = tr.dataset.img;
                  if (!id) return;

                  tr.style.backgroundColor = '#eee'; // tô xám dòng

                  const sideBox = document.getElementById('side-image-box');
                  const sideImg = document.getElementById('side-image');
                  sideImg.src = `${inputValue}/${id}`;
                  sideBox.style.display = 'block';
                });

                tr.addEventListener('mouseleave', () => {
                  tr.style.backgroundColor = ''; // bỏ tô xám khi rời chuột

                  const sideBox = document.getElementById('side-image-box');
                  sideBox.style.display = 'none';
                });
              });
            // Gắn click cho chữ ở cột E để sửa chữ ở cột C
            document.querySelectorAll('.char-ex').forEach(span => {
              span.addEventListener('click', () => {
                const originalText = span.dataset.char;
                const row = span.dataset.row;
                const index = span.dataset.index;

                const cxSpan = document.querySelector(`.char-cx[data-row="${row}"][data-index="${index}"]`);
                if (cxSpan) {
                  cxSpan.style.backgroundColor = '#ffff99'; // tô vàng khi chọn
                }
                fetch(`http://127.0.0.1:5000/suggest?char=${encodeURIComponent(originalText)}`)
                  .then(r => r.json())
                  .then(data => {
                    const suggestions = data.suggestions;
                    if (!suggestions.length) return alert(`Không có gợi ý cho '${originalText}'`);

                  const select = document.createElement('select');

                  // Thêm option rỗng đầu tiên
                  const emptyOption = document.createElement('option');
                  emptyOption.textContent = ''; // hiển thị khoảng trắng
                  emptyOption.disabled = true;
                  emptyOption.selected = true;
                  select.appendChild(emptyOption);

                  // Thêm các gợi ý từ API
                  suggestions.forEach(s => {
                    const option = document.createElement('option');
                    option.value = s;
                    option.textContent = s;
                    select.appendChild(option);
                  });


                  select.addEventListener('change', () => {
                    applyChange();
                  });

                  // Cũng gọi khi vừa tạo dropdown — nếu người không đổi, vẫn cập nhật
                  select.addEventListener('blur', () => {
                    if (document.body.contains(select)) {
                      applyChange();
                    }
                  });

                  function applyChange() {
                    const newChar = select.value;
                    const cxSpan = document.querySelector(`.char-cx[data-row="${row}"][data-index="${index}"]`);
                    if (cxSpan) {
                      cxSpan.textContent = newChar;
                      cxSpan.style.color = 'green';
                      cxSpan.style.fontWeight = 'bold';
                      cxSpan.style.backgroundColor = '';
                    }
                    select.remove();
                  }

                    span.after(select);
                  })
                  .catch(err => {
                    console.error('❌ Lỗi gợi ý:', err);
                    alert(`Không thể lấy gợi ý cho '${originalText}'`);
                  });
              });
            });
// Gắn hover highlight đồng thời cho cả char-ex và char-cx
document.querySelectorAll('.char-ex').forEach(span => {
  span.addEventListener('mouseenter', () => {
    const row = span.dataset.row;
    const index = span.dataset.index;

    const cxSpan = document.querySelector(`.char-cx[data-row="${row}"][data-index="${index}"]`);
    if (cxSpan) {
      cxSpan.style.backgroundColor = '#ffff99';
      span.style.backgroundColor = '#ffff99';  // vàng nhạt
      cxSpan.style.backgroundColor = '#ffff99';
    }
  });

  span.addEventListener('mouseleave', () => {
    const row = span.dataset.row;
    const index = span.dataset.index;

    const cxSpan = document.querySelector(`.char-cx[data-row="${row}"][data-index="${index}"]`);
    if (cxSpan) {
      span.style.backgroundColor = '';
      cxSpan.style.backgroundColor = '';
    }
  });
});

document.querySelectorAll('.char-cx').forEach(span => {
  span.addEventListener('mouseenter', () => {
    const row = span.dataset.row;
    const index = span.dataset.index;

    const exSpan = document.querySelector(`.char-ex[data-row="${row}"][data-index="${index}"]`);
    if (exSpan) {
      span.style.backgroundColor = '#ffffcc';  // khác màu tí
      exSpan.style.backgroundColor = '#ffffcc';
    }
  });

  span.addEventListener('mouseleave', () => {
    const row = span.dataset.row;
    const index = span.dataset.index;

    const exSpan = document.querySelector(`.char-ex[data-row="${row}"][data-index="${index}"]`);
    if (exSpan) {
      span.style.backgroundColor = '';
      exSpan.style.backgroundColor = '';
    }
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

function inlineAllStyles(element) {
  const clone = element.cloneNode(true);
  const allElements = clone.querySelectorAll("*");

  allElements.forEach(el => {
    const computedStyle = getComputedStyle(el);
    for (let i = 0; i < computedStyle.length; i++) {
      const key = computedStyle[i];
      el.style[key] = computedStyle.getPropertyValue(key);
    }
  });

  const rootStyle = getComputedStyle(clone);
  for (let i = 0; i < rootStyle.length; i++) {
    const key = rootStyle[i];
    clone.style[key] = rootStyle.getPropertyValue(key);
  }

  return clone.outerHTML;
}

document.getElementById('send-table-to-server').addEventListener('click', () => {
  const outputDiv = document.getElementById('output');
  const path_excel = document.getElementById("path_excel").value;
  const table = outputDiv.querySelector('table');

  if (!table) {
    alert("Không tìm thấy bảng!");
    return;
  }

  const inlineHTML = inlineAllStyles(table);

  fetch('http://127.0.0.1:5000/save_table', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      table_html: inlineHTML,
      path_excel: path_excel // ✅ thêm dòng này
    })
  })
  .then(res => res.json())
  .then(data => alert("✅ Đã gửi bảng về server"))
  .catch(err => {
    console.error("❌ Lỗi khi gửi bảng:", err);
    alert("Không gửi được bảng về server");
  });
});
