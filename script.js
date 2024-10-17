const codigoMap = {
    'OBGP': { 6: 87, 7: 87, 8: 89, 9: 89, 1: 93, 2: 93, 3: 93, 0: 91 },
    'OBLI': { 4: 94, 5: 94, 6: 96, 7: 96, 8: 98, 9: 98, 1: 102, 2: 102, 3: 102, 0: 100 },
    'OBMF': { 4: 104, 5: 104, 6: 106, 7: 106, 8: 108, 9: 108, 1: 110, 2: 110, 3: 110, 0: 112 }
};

document.getElementById('process-button').addEventListener('click', () => {
    const fileInput = document.getElementById('file-input');
    const file = fileInput.files[0];

    if (!file) {
        alert('Por favor, selecione um arquivo.');
        return;
    }

    const reader = new FileReader();
    reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        processData(jsonData);
    };
    reader.readAsArrayBuffer(file);
});

let processedData = [];

function processData(data) {
    processedData = [];
    let errors = []; // Lista para armazenar erros
    let studentCount = 0; // Contador de alunos

    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (row.length < 4) {
            console.warn(`Linha ${i + 1} incompleta: ${row}`);
            continue;
        }

        const nome = row[1] ? row[1].toString().trimEnd() : '';  
        const email = row[2] ? row[2].toString().trim().replace(/\s+/g, '') : '';  
        const olympicAbbreviation = row[3] ? row[3].toString().trim() : '';  
        const grade = (row[4] !== undefined && row[4] !== null) ? parseInt(row[4].toString().trim()) : null; 

        const cleanedNome = nome.replace(/[^a-zA-ZÀ-ÿ\sãâá]/g, '').replace(/[^a-zA-ZÀ-ÿ\s]/g, ""); 
        const normalizedNome = cleanedNome.normalize('NFD').replace(/[\u0300-\u036f]/g, "");
        const cleanedEmail = email.normalize('NFD').replace(/[\u0300-\u036f]/g, "").replace(/[^a-zA-Z0-9@._-]/g, '');

        if (codigoMap[olympicAbbreviation] && codigoMap[olympicAbbreviation].hasOwnProperty(grade)) {
            const code = codigoMap[olympicAbbreviation][grade];
            if (normalizedNome && cleanedEmail) {
                processedData.push([normalizedNome, cleanedEmail, code]);
                studentCount++; // Incrementa o contador de alunos
            }
        } else {
            errors.push(`Erro na linha ${i + 1}: Olimpíada - ${olympicAbbreviation}, Série - ${grade}`); // Armazena o erro
        }
    }

    // Atualiza o contador de alunos
    document.getElementById('student-count').innerText = studentCount;

    // Exibe erros ou mensagem de "Nenhum erro encontrado"
    const errorLog = document.getElementById('error-log');
    if (errors.length > 0) {
        errorLog.innerText = errors.join('\n'); // Exibe os erros
    } else {
        errorLog.innerText = "Nenhum erro encontrado"; // Exibe mensagem padrão
    }

    // Dividir em arquivos de até 100 alunos
    const files = [];
    for (let i = 0; i < processedData.length; i += 100) {
        const chunk = processedData.slice(i, i + 100);
        files.push(chunk);
    }

    if (files.length > 0) {
        const outputDiv = document.getElementById('output');
        outputDiv.innerHTML = files.map((file, index) => {
            return `<h3>Arquivo ${index + 1}</h3><pre>${file.map(row => row.join(', ')).join('\n')}</pre>`;
        }).join('');
        document.getElementById('output-container').style.display = 'block';
        document.getElementById('send-button').style.display = 'inline-block';
        document.getElementById('download-button').style.display = 'inline-block';
    } else {
        alert('Nenhum dado válido encontrado.');
    }
}



document.getElementById('download-button').addEventListener('click', () => {
    downloadCSV(processedData);
});

function downloadCSV(data) {
    // Dividindo arquivos
    const files = [];
    for (let i = 0; i < data.length; i += 100) {
        files.push(data.slice(i, i + 100));
    }

    files.forEach((fileData, index) => {
        let csvContent = fileData.map(rowArray => rowArray.join(",")).join("\r\n");
        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement("a");
        const url = URL.createObjectURL(blob);
        link.setAttribute("href", url);
        link.setAttribute("download", `arquivo_processado_${index + 1}.csv`);
        link.style.visibility = 'hidden';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    });
}
