document.addEventListener('DOMContentLoaded', () => {
    const excelFile = document.getElementById('excelFile');
    const readButton = document.getElementById('readButton');
    const tableSection = document.getElementById('tableSection');
    const dataTable = document.getElementById('dataTable');
    const generatePdfButton = document.getElementById('generatePdfButton');
    const messageArea = document.getElementById('messageArea');

    // Novos elementos para o filtro de datas
    const startDateInput = document.getElementById('startDate');
    const endDateInput = document.getElementById('endDate');
    const applyFilterButton = document.getElementById('applyFilterButton');

    const hiddenImageContainer = document.createElement('div');
    hiddenImageContainer.style.position = 'absolute';
    hiddenImageContainer.style.left = '-9999px';
    hiddenImageContainer.style.top = '-9999px';
    hiddenImageContainer.style.width = '1px';
    hiddenImageContainer.style.height = '1px';
    hiddenImageContainer.style.overflow = 'hidden';
    document.body.appendChild(hiddenImageContainer);

    let originalExcelData = []; // Armazena os dados brutos da planilha
    let filteredExcelData = []; // Armazena os dados filtrados

    readButton.addEventListener('click', () => {
        const file = excelFile.files[0];
        if (!file) {
            showMessage('Por favor, selecione um arquivo Excel.', 'warning');
            return;
        }

        const reader = new FileReader();

        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];

            originalExcelData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            if (originalExcelData.length > 0) {
                // Ao carregar a planilha, inicialmente os dados filtrados são os originais
                filteredExcelData = [...originalExcelData]; // Cria uma cópia para evitar modificação direta
                renderTable(filteredExcelData);
                tableSection.classList.remove('hidden');
                showMessage('Planilha lida com sucesso!', 'success');
            } else {
                showMessage('A planilha está vazia ou não pôde ser lida.', 'warning');
                tableSection.classList.add('hidden');
            }
        };

        reader.onerror = (e) => {
            showMessage('Erro ao ler o arquivo: ' + e.target.error.name, 'error');
            tableSection.classList.add('hidden');
        };

        reader.readAsArrayBuffer(file);
    });

    // Event listener para o botão de aplicar filtro
    applyFilterButton.addEventListener('click', () => {
        if (originalExcelData.length === 0) {
            showMessage('Nenhum dado carregado para filtrar. Por favor, leia um arquivo Excel primeiro.', 'warning');
            return;
        }

        const startDate = startDateInput.value ? new Date(startDateInput.value) : null;
        const endDate = endDateInput.value ? new Date(endDateInput.value) : null;

        if (!startDate && !endDate) {
            // Se nenhum filtro de data for selecionado, mostre todos os dados originais
            filteredExcelData = [...originalExcelData];
            renderTable(filteredExcelData);
            showMessage('Nenhum filtro de data aplicado. Exibindo todos os dados.', 'info');
            return;
        }

        const headers = originalExcelData[0];
        const dateColIndex = headers.findIndex(h => h && typeof h === 'string' && h.toLowerCase().includes('carimbo de data/hora'));

        if (dateColIndex === -1) {
            showMessage('Coluna "Carimbo de data/hora" não encontrada para filtrar.', 'error');
            return;
        }

        filteredExcelData = originalExcelData.filter((row, index) => {
            if (index === 0) return true; // Sempre incluir o cabeçalho

            const rowDateRaw = row[dateColIndex];
            if (!rowDateRaw) return false;

            // Tenta parsear a data. O XLSX.read pode retornar datas como números ou strings.
            // Para datas como números (formato Excel), use XLSX.SSF.parse_date_code.
            // Para strings, tente new Date().
            let cellDate;
            if (typeof rowDateRaw === 'number') {
                // Assumindo que a coluna de data esteja no formato numérico do Excel
                const dateObj = XLSX.SSF.parse_date_code(rowDateRaw);
                cellDate = new Date(dateObj.y, dateObj.m - 1, dateObj.d, dateObj.H, dateObj.M, dateObj.S);
            } else {
                cellDate = new Date(rowDateRaw);
            }
            
            // Valida a data para garantir que é uma data válida antes de comparar
            if (isNaN(cellDate.getTime())) {
                console.warn(`Data inválida encontrada na linha ${index + 1}: ${rowDateRaw}. Esta linha será ignorada no filtro.`);
                return false; 
            }

            // Define o início do dia para as datas de filtro para comparar apenas o dia
            const filterStartDate = startDate ? new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate()) : null;
            const filterEndDate = endDate ? new Date(endDate.getFullYear(), endDate.getMonth(), endDate.getDate(), 23, 59, 59, 999) : null; // Fim do dia

            const isAfterStartDate = !filterStartDate || cellDate >= filterStartDate;
            const isBeforeEndDate = !filterEndDate || cellDate <= filterEndDate;

            return isAfterStartDate && isBeforeEndDate;
        });

        // Se após o filtro, restar apenas o cabeçalho, significa que não há dados correspondentes
        if (filteredExcelData.length === 1 && originalExcelData.length > 1) {
            showMessage('Nenhum dado encontrado para as datas selecionadas.', 'info');
        } else if (filteredExcelData.length > 0) {
            showMessage(`Filtro aplicado! Exibindo ${filteredExcelData.length - 1} linhas.`, 'success');
        }

        renderTable(filteredExcelData);
    });

    // Função para transformar URLs do Google Drive em links de miniatura
    function getGoogleDriveThumbnailUrl(shareUrl, size = 1000) {
        const match = shareUrl.match(/https?:\/\/drive\.google\.com\/(?:file\/d\/|open\?id=|uc\?id=|thumbnail\?id=)([a-zA-Z0-9_-]+)/);
        
        if (match && match[1]) {
            const fileId = match[1];
            // Esta URL 'https://lh3.googleusercontent.com/d/${fileId}=w${size}' parece incorreta para thumbnails do Google Drive.
            // A URL correta para thumbnails geralmente é: `https://drive.google.com/thumbnail?id=${fileId}&sz=w${size}`
            // Ou para visualização direta, pode ser `https://drive.google.com/uc?id=${fileId}&export=download`
            // Ou para uma imagem mais robusta, `https://lh3.googleusercontent.com/d/${fileId}=w${size}` (se a imagem foi enviada via um formulário do Google Forms por exemplo)
            // Vou manter a sua, mas tenha em mente que pode não funcionar para todos os casos de uso.
            return `https://lh3.googleusercontent.com/d/${fileId}=w${size}`; 
        }
        return shareUrl; 
    }

    function loadImage(url) {
        return new Promise((resolve, reject) => {
            const img = new Image();
            img.crossOrigin = 'Anonymous';

            img.onload = () => {
                const canvas = document.createElement('canvas');
                canvas.width = img.width;
                canvas.height = img.height;
                const ctx = canvas.getContext('2d');
                ctx.drawImage(img, 0, 0);

                try {
                    const dataUrl = canvas.toDataURL('image/jpeg');
                    resolve({ dataUrl, width: img.width, height: img.height });
                } catch (e) {
                    reject(new Error(`Erro ao gerar dataURL: ${e.message}`));
                }
            };

            img.onerror = () => {
                reject(new Error(`Erro ao carregar imagem: ${url}`));
            };

            img.src = url;
        });
    }

    generatePdfButton.addEventListener('click', async () => {
        // Agora usa filteredExcelData para gerar o PDF
        if (filteredExcelData.length < 2) {
            showMessage('Nenhum dado filtrado para gerar PDF. Por favor, leia um arquivo com conteúdo e aplique um filtro se necessário.', 'warning');
            return;
        }

        showMessage('Gerando PDF... Isso pode demorar para muitas imagens. Não feche a página.', 'info');

        const { jsPDF } = window.jspdf;
        const doc = new jsPDF('portrait', 'mm', 'a4');

        const headers = filteredExcelData[0];
        const dataRows = filteredExcelData.slice(1); // Usa filteredExcelData.slice(1)

        const dateColIndex = headers.findIndex(h => h && typeof h === 'string' && h.toLowerCase().includes('carimbo de data/hora'));
        const photoColIndex = headers.findIndex(h => h && typeof h === 'string' && h.toLowerCase().includes('foto'));
        const authorColIndex= headers.findIndex(h => h && typeof h === 'string' && h.toLowerCase().includes('nome completo'));
        const sectorColIndex = headers.findIndex(h => h && typeof h === 'string' && h.toLowerCase().includes('setor'));
        const descriptionColIndex = headers.findIndex(h => h && typeof h === 'string' && h.toLowerCase().includes('descrição'));

        if (dateColIndex === -1 || photoColIndex === -1 || descriptionColIndex === -1) {
            showMessage('As colunas "Carimbo de data/hora", "foto" ou "descrição" não foram encontradas na planilha. Verifique os nomes exatos das colunas.', 'error');
            return;
        }

        let yOffset = 15;
        const margin = 10;
        const lineHeight = 6;
        const imageMaxWidth = doc.internal.pageSize.width - (margin * 2);
        const maxPageHeight = doc.internal.pageSize.height - (margin * 2);

        const checkPageBreak = (requiredSpace) => {
            if (yOffset + requiredSpace > maxPageHeight) {
                doc.addPage();
                yOffset = margin;
            }
        };


        doc.setFontSize(18);
        doc.setFont('helvetica', 'bold');
        doc.text('Relátorio de Inspeção UMB', margin, yOffset);
        yOffset += 2*lineHeight;

        for (const row of dataRows) {
            // const date = String(row[dateColIndex] || 'N/A');
            const date = formatarDataHoraCompleta(row[dateColIndex]);
            const photosString = String(row[photoColIndex] || 'Não Informado');
            const author = String(row[authorColIndex] || 'Não Informado');
            const sector = String(row[sectorColIndex] || 'Não Informado');
            const description = String(row[descriptionColIndex] || 'N/A');

            checkPageBreak(lineHeight);
            doc.setFontSize(11);
            doc.setFont('helvetica', 'bold');
            doc.text(`Autor: ${author}`, margin, yOffset);
            yOffset += lineHeight;
            doc.setFontSize(9);
            doc.setFont('helvetica', 'bold');
            doc.text(`Setor: ${sector}`, margin, yOffset);
            yOffset += lineHeight;
            doc.text(`Data/Hora: ${date}`, margin, yOffset);
            yOffset += lineHeight;

            doc.setFontSize(9);
            doc.setFont('helvetica', 'normal');
            const splitDescription = doc.splitTextToSize(`Descrição: ${description}`, doc.internal.pageSize.width - margin * 2);
            checkPageBreak(splitDescription.length * lineHeight);
            doc.text(splitDescription, margin, yOffset);
            yOffset += (splitDescription.length * lineHeight) + 2;

            const photoUrls = photosString.split(',').map(url => url.trim()).filter(url => url !== '');

            if (photoUrls.length > 0) {
                doc.setFontSize(9);
                doc.setFont('helvetica', 'italic');
                checkPageBreak(lineHeight);
                doc.text('Imagens:', margin, yOffset);
                yOffset += lineHeight + 2;

                for (const url of photoUrls) {
                    const processedUrl = getGoogleDriveThumbnailUrl(url, 1000); 
                    try {
                        const { dataUrl, width, height } = await loadImage(processedUrl);

                        // let imgDisplayWidth = imageMaxWidth;
                        // let imgDisplayHeight = (height * imgDisplayWidth) / width;

                        let imgDisplayWidth = 150;
                        let imgDisplayHeight = 100; 

                        if (imgDisplayHeight > (maxPageHeight - yOffset - 5)) {
                            imgDisplayHeight = maxPageHeight - yOffset - 5;
                            imgDisplayWidth = (width * imgDisplayHeight) / height;
                        }

                        checkPageBreak(imgDisplayHeight + 5);
                        const imgX = margin + (imageMaxWidth - imgDisplayWidth) / 2;
                        doc.addImage(dataUrl, 'JPEG', imgX, yOffset, imgDisplayWidth, imgDisplayHeight);

                        doc.link(imgX, yOffset, imgDisplayWidth, imgDisplayHeight, {
                            url: processedUrl // link destination
                        });

                        yOffset += imgDisplayHeight + 5;

                    } catch (error) {
                        console.error(`Erro ao carregar imagem da URL: ${processedUrl}`, error);
                        checkPageBreak(lineHeight + 5);
                        doc.setFontSize(8);
                        doc.setFont('helvetica', 'normal');
                        doc.setTextColor(255, 0, 0);
                        doc.text(`[ERRO] Não foi possível carregar imagem: ${processedUrl}`, margin, yOffset);
                        doc.setTextColor(0, 0, 0);
                        yOffset += lineHeight + 5;
                    }
                }
            } else {
                checkPageBreak(lineHeight);
                doc.setFontSize(9);
                doc.setFont('helvetica', 'normal');
                doc.text('Imagens: Nenhuma imagem fornecida.', margin, yOffset);
                yOffset += lineHeight;
            }

            yOffset += 10;
            checkPageBreak(20);
            doc.setDrawColor(200, 200, 200);
            doc.line(margin, yOffset - 5, doc.internal.pageSize.width - margin, yOffset - 5);
        }

        const pageCount = doc.internal.getNumberOfPages();
        for (let i = 1; i <= pageCount; i++) {
            doc.setPage(i);
            doc.setFontSize(8);
            doc.setFont('helvetica', 'normal');
            doc.text(`Página ${i} de ${pageCount}`, doc.internal.pageSize.width - margin, doc.internal.pageSize.height - margin, { align: 'right' });
        }

        doc.save('dados_personalizados.pdf');
        showMessage('PDF gerado com sucesso!', 'success');
    });

    // Função para renderizar a tabela, agora recebe os dados a serem exibidos
    function renderTable(dataToRender) {
        const thead = dataTable.querySelector('thead');
        const tbody = dataTable.querySelector('tbody');

        thead.innerHTML = '';
        tbody.innerHTML = '';

        if (dataToRender.length > 0) {
            const headerRow = document.createElement('tr');
            dataToRender[0].forEach(headerText => {
                const th = document.createElement('th');
                th.textContent = headerText;
                headerRow.appendChild(th);
            });
            thead.appendChild(headerRow);

            for (let i = 1; i < dataToRender.length; i++) {
                const rowData = dataToRender[i];
                const tr = document.createElement('tr');
                const dateColIndex = dataToRender[0].findIndex(h => h && typeof h === 'string' && h.toLowerCase().includes('carimbo de data/hora'));

                rowData.forEach((cellData, colIndex) => {
                    const td = document.createElement('td');
                    if (colIndex === dateColIndex) {
                        // --- AQUI A DATA É FORMATADA PARA A TABELA HTML ---
                        td.textContent = formatarDataHoraCompleta(cellData);
                        // --- FIM DA FORMATAÇÃO ---
                    } else {
                        td.textContent = cellData;
                    }
                    tr.appendChild(td);
                });
                tbody.appendChild(tr);
            }
        }
    }

    function showMessage(message, type = 'info') {
        messageArea.textContent = message;
        messageArea.className = 'message-area show';
        messageArea.style.backgroundColor = '';
        messageArea.style.color = '';

        if (type === 'success') {
            messageArea.style.backgroundColor = '#d4edda';
            messageArea.style.color = '#155724';
        } else if (type === 'warning') {
            messageArea.style.backgroundColor = '#fff3cd';
            messageArea.style.color = '#856404';
        } else if (type === 'error') {
            messageArea.style.backgroundColor = '#f8d7da';
            messageArea.style.color = '#721c24';
        } else if (type === 'info') {
            messageArea.style.backgroundColor = '#cce5ff';
            messageArea.style.color = '#004085';
        }

        setTimeout(() => {
            messageArea.classList.remove('show');
        }, 5000);
    }

    function formatarDataHoraCompleta(dataRaw) {
        if (!dataRaw) return 'N/A';

        let dataObj;

        // Attempt to parse the date. XLSX.read can return dates as numbers or strings.
        if (typeof dataRaw === 'number') {
            // If it's a number, assume it's an Excel date format
            // Note: For XLSX.SSF.parse_date_code to work, you'd typically need the XLSX library
            // loaded in your environment. If you're not using XLSX, this part might need adjustment
            // or removal if your numbers are standard Unix timestamps (milliseconds since epoch).
            try {
                // This line assumes XLSX.SSF is available globally.
                // If not, and 'dataRaw' is a Unix timestamp (milliseconds), use:
                // dataObj = new Date(dataRaw);
                const excelDate = XLSX.SSF.parse_date_code(dataRaw);
                dataObj = new Date(excelDate.y, excelDate.m - 1, excelDate.d, excelDate.H, excelDate.M, excelDate.S);
            } catch (e) {
                // Fallback for when XLSX.SSF is not available or it's a standard timestamp
                dataObj = new Date(dataRaw);
            }
        } else {
            // If it's a string, try to create a Date object directly
            dataObj = new Date(dataRaw);
        }

        // Check if the date is valid
        if (isNaN(dataObj.getTime())) {
            return String(dataRaw); // Return the original value if it's not a valid date
        }

        const dia = String(dataObj.getDate()).padStart(2, '0');
        const mes = String(dataObj.getMonth() + 1).padStart(2, '0'); // Month is 0-11
        const ano = dataObj.getFullYear();
        const hora = String(dataObj.getHours()).padStart(2, '0');
        const minuto = String(dataObj.getMinutes()).padStart(2, '0');
        const segundo = String(dataObj.getSeconds()).padStart(2, '0');

        return `${dia}/${mes}/${ano} ${hora}:${minuto}:${segundo}`;
    }
});
