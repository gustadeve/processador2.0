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
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (row.length < 4) {
            console.warn(`Linha ${i + 1} incompleta: ${row}`);
            continue;
        }

        const nome = row[1] ? row[1].toString().trimEnd() : '';  
        const email = row[2] ? row[2].toString().trim().replace(/\s+/g, '') : '';  
        const olympicAbbreviation = row[3] ? row[3].toString().trim() : '';  
        const grade = row[4] ? parseInt(row[4].toString().trim()) : null;

        // Remove números do nome, mantendo apenas letras e acentos agudos e til
        const cleanedNome = nome.replace(/[^a-zA-ZÀ-ÿ\sãâá]/g, '').replace(/[^a-zA-ZÀ-ÿ\s]/g, ""); 

        // Remove acentos do nome
        const normalizedNome = cleanedNome.normalize('NFD').replace(/[\u0300-\u036f]/g, "");

        // Limpa o email, removendo caracteres não permitidos e acentos
        const cleanedEmail = email.normalize('NFD').replace(/[\u0300-\u036f]/g, "").replace(/[^a-zA-Z0-9@._-]/g, '');

        if (codigoMap[olympicAbbreviation] && codigoMap[olympicAbbreviation][grade] !== undefined) {
            const code = codigoMap[olympicAbbreviation][grade];
            if (normalizedNome && cleanedEmail) {
                processedData.push([normalizedNome, cleanedEmail, code]);
            }
        } else {
            console.warn(`Dados inválidos na linha ${i + 1}: Olimpíada - ${olympicAbbreviation}, Série - ${grade}`);
        }
    }

    if (processedData.length > 0) {
        const outputDiv = document.getElementById('output');
        outputDiv.innerHTML = processedData.map(row => row.join(', ')).join('\n');
        document.getElementById('output-container').style.display = 'block';
        document.getElementById('send-button').style.display = 'inline-block';
        document.getElementById('download-button').style.display = 'inline-block';
    } else {
        alert('Nenhum dado válido encontrado.');
    }
}


document.getElementById('send-button').addEventListener('click', () => {
    // Gerar o CSV
    const csvContent = processedData.map(rowArray => rowArray.join(",")).join("\r\n");
    
    // Função de upload
    function uploadFile() {
        const inputFile = document.getElementById('import_file');
        const file = new Blob([csvContent], { type: 'text/csv' });
        const dataTransfer = new DataTransfer();
        dataTransfer.items.add(new File([file], 'alunos_processados.csv'));
        inputFile.files = dataTransfer.files;

        // Dispara um evento de mudança no input
        const event = new Event('change', { bubbles: true });
        inputFile.dispatchEvent(event);
    }

    // Chama a função para fazer o upload
    uploadFile();

    // Alerta de sucesso
    alert('Arquivo enviado com sucesso!');
});

document.getElementById('download-button').addEventListener('click', () => {
    downloadCSV(processedData);
});

function downloadCSV(data) {
    let csvContent = data.map(rowArray => rowArray.join(",")).join("\r\n");
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement("a");
    const url = URL.createObjectURL(blob);
    link.setAttribute("href", url);
    link.setAttribute("download", "arquivo_processado.csv");
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}
document.getElementById('send-button').addEventListener('click', () => {
    const fileInput = document.getElementById('import_file');
    const file = new Blob([csvContent], { type: 'text/csv' }); // onde csvContent é a string CSV gerada.
    const dataTransfer = new DataTransfer();
    dataTransfer.items.add(new File([file], 'arquivo_processado.csv'));

    fileInput.files = dataTransfer.files; // Adiciona o arquivo ao campo de input

    // Agora simula o envio do formulário se necessário
    const form = fileInput.closest('form'); // Isso assume que o campo de input está dentro de um form
    if (form) {
        form.submit(); // Isso submeterá o formulário
    }
});