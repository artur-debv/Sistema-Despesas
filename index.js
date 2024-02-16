document.addEventListener('DOMContentLoaded', function () {
    let wb = null;
  
    function exportToExcel() {
      if (!wb) {
        wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet([
          ['Nome do Produto', 'Tipo de Conta', 'Data', 'Valor'], // Defina valores com estilo
        ], {
          cellStyles: {
            A1: {
              font: { bold: true },
              fill: { fgColor: { rgb: 'FF0000' } }, // Vermelho de fundo
              border: { top: { style: 'thin' }, bottom: { style: 'thin' } }
            },
            B1: {
              font: { bold: true },
              fill: { fgColor: { rgb: 'FF0000' } }, // Vermelho de fundo
              border: { top: { style: 'thin' }, bottom: { style: 'thin' } }
            },
            C1: {
              font: { bold: true },
              fill: { fgColor: { rgb: 'FF0000' } }, // Vermelho de fundo
              border: { top: { style: 'thin' }, bottom: { style: 'thin' } }
            },
            D1: {
              font: { bold: true },
              fill: { fgColor: { rgb: 'FF0000' } }, // Vermelho de fundo
              border: { top: { style: 'thin' }, bottom: { style: 'thin' } }
            }
          }
        });
        XLSX.utils.book_append_sheet(wb, ws, 'Despesas');
      }
  
      const nomeProduto = document.getElementById('nome-produto').value;
      const tipoConta = document.getElementById('tipo-conta').value;
      const data = document.getElementById('data').value;
      const valor = document.getElementById('valor').value;
  
      const ws = wb.Sheets['Despesas'];
      const lastRow = XLSX.utils.decode_range(ws['!ref']).e.r + 1;
      const newValues = [nomeProduto, tipoConta, data, valor];
  
      newValues[2] = { v: data, t: 'd', z: 'mm/dd/yyyy' }; // Formata data como data curta
      newValues[3] = { v: valor, t: 'n', z: '#,##0.00' }; // Formata valor como moeda
  
      XLSX.utils.sheet_add_aoa(ws, [newValues], {
        origin: -1,
        cellStyles: {
          A2: {
            border: {
              top: { style: 'thin' },
              right: { style: 'thin' },
              bottom: { style: 'thin' },
              left: { style: 'thin' }
            }
          },
          B2: {
            border: {
              top: { style: 'thin' },
              right: { style: 'thin' },
              bottom: { style: 'thin' },
              left: { style: 'thin' }
            }
          },
          C2: {
            border: {
              top: { style: 'thin' },
              right: { style: 'thin' },
              bottom: { style: 'thin' },
              left: { style: 'thin' }
            }
          },
          D2: {
            border: {
              top: { style: 'thin' },
              right: { style: 'thin' },
              bottom: { style: 'thin' },
              left: { style: 'thin' }
            }
          }
        }
      });
  
      ws['!cols'] = [{ width: 20 }, { width: 20 }, { width: 15 }, { width: 15 }]; // Largura personalizada para cada coluna
  
      // Exporta o arquivo Excel
      XLSX.writeFile(wb, 'despesas.xlsx');
    }
  
    document.getElementById('form-despesa').addEventListener('submit', function (event) {
      event.preventDefault();
      exportToExcel();
    });
  });