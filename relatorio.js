function loadFile(url, callback) {
    PizZipUtils.getBinaryContent(url, callback);
}

function gerar(alunos, curso) {
    loadFile(
        "https://raw.githubusercontent.com/jrobertogram/cvs/main/modeloJS.docx",
        function(error, content) {
            if (error) {
                throw error;
            }
            const zip = new PizZip(content);
            const doc = new window.docxtemplater(zip, {
                paragraphLoop: true,
                linebreaks: true,
            });

            let data = getDataAtualFormatada()

            alunosDoc = [];
            alunos.forEach(function(element, index) {
                alunosDic = {};
                alunosDic['nome'] = element[2];
                alunosDic['email'] = element[1];
                alunosDic['matricula'] = element[4];
                alunosDic['cpf'] = element[3];
                alunosDoc.push(alunosDic);
            });

            doc.render({
                "curso": curso['curso'],
                "monitor": curso['monitor'],
                "ch": curso['ch'],
                "datas": curso['datas'],
                "data": data,
                "alunos": alunosDoc
            });

            const blob = doc.getZip().generate({
                type: "blob",
                mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                compression: "DEFLATE",
            });
            saveAs(blob, "saida.docx");
        }
    );
};

function criarRelatorio(callback) {
    fazerPresenca(function(alunos_presentes) {
        if (alunos_presentes !== null) {

            curso = {}
               
            curso['curso'] = document.getElementById("nome").value;
            curso['monitor'] = document.getElementById("monitor").value;
            curso['ch'] = document.getElementById("ch").value;
            curso['datas'] = document.getElementById("datas").value;
            gerar(alunos_presentes, curso)
            //console.log(alunos_presentes);

        } else {
            console.error('Ocorreu um erro ao processar a presença dos alunos.');
        }
    });
}

function fazerPresenca(callback) {
    JuntarCsv(DropzonePresenca.files, function(cpfs_presenca_marge) {
        if (cpfs_presenca_marge !== null) {
            presenca = document.getElementById("presenca").value  ?? 0;
            cpfs_presente = filtrarPresenca(cpfs_presenca_marge, presenca);
            filtrarAlunosPresentes(cpfs_presente, function(alunos_presentes) {
                if (alunos_presentes !== null) {
                    callback(alunos_presentes);
                } else {
                    console.error('Ocorreu um erro ao processar o arquivo.');
                    callback(null);
                }
            });
        } else {
            console.error('Ocorreu um erro ao processar os arquivos.');
            callback(null);
        }
    });
}


function filtrarAlunosPresentes(cpfs_presente, callback) {

    if (DropzoneAlunos.files.length === 0) {
        console.error('Nenhum arquivo selecionado.');
        callback(null);
        return;
    }

    const file = DropzoneAlunos.files[0];
    const reader = new FileReader();

    reader.onload = function(event) {
        const csvData = event.target.result;
        const alunos_presentes = [];

        const csvLines = csvData.split('\n');
        csvLines.forEach(line => {
            line = line.replace(/"/g, '');
            const alunoData = line.split(',');
            const cpf_aluno = alunoData[3];

            if (cpfs_presente.includes((cpf_aluno))) {
                alunos_presentes.push(alunoData);
            }
        });

        callback(alunos_presentes);
    };

    reader.onerror = function(event) {
        console.error('Ocorreu um erro ao ler o arquivo:', event.target.error);
        callback(null);
    };

    reader.readAsText(file);
}

function filtrarPresenca(cpfs_presenca, quantidadePresente) {
    const presente = cpfs_presenca.reduce(function(obj, cpf) {
        obj[cpf] = obj[cpf] ? obj[cpf] + 1 : 1;
        return obj;
    }, {});

    const qualificado = Object.keys(presente).filter(function(cpf) {
        return presente[cpf] >= quantidadePresente;
    });

    return qualificado;
}



function JuntarCsv(files, callback) {
    if (files && files.length > 0) {
        const cpfs_presenca_marge = [];

        let filesProcessed = 0;

        function processFile(file) {
            const reader = new FileReader();

            reader.onload = function(e) {
                const csvData = e.target.result;
                const csvLines = csvData.split('\n');
                csvLines.shift();

                for (let j = 0; j < csvLines.length; j++) {
                    const line = csvLines[j].replace(/"/g, '');
                    const cpf = line.split(',')[2];
                    cpfs_presenca_marge.push(cpf);
                }

                filesProcessed++;

                if (filesProcessed === files.length) {
                    callback(cpfs_presenca_marge);
                }
            };

            reader.onerror = function(e) {
                console.error('Ocorreu um erro ao ler o arquivo:', e.target.error);
                callback(null);
            };

            reader.readAsText(file);
        }

        for (let i = 0; i < files.length; i++) {
            const file = files[i];
            processFile(file);
        }
    } else {
        console.error('Nenhum arquivo selecionado.');
        callback(null);
    }
}

function getDataAtualFormatada() {
    moment.locale('pt');
    let dataAtual = moment();

    let nomeMes = dataAtual.format("MMMM");
    nomeMes = nomeMes.charAt(0).toUpperCase() + nomeMes.slice(1);

    let dataFormatada = dataAtual.format("DD [de] ") + nomeMes + dataAtual.format(" [de] YYYY");
    return dataFormatada;
}

Dropzone.autoDiscover = false;

let configDropzone = {
    addRemoveLinks: true,
    acceptedFiles: "text/csv",
    dictRemoveFile: "Remover arquivo",
    dictFallbackMessage: "Seu navegador não suporta arrastar e soltar para fazer upload de arquivos.",
    dictFileTooBig: "O arquivo é muito grande ({{filesize}} MB). Tamanho máximo do arquivo: {{maxFilesize}} MB.",
    dictInvalidFileType: "Tipo de arquivo inválido. Apenas arquivos CSV são permitidos.",
}

configDropzone['dictDefaultMessage'] = "Arraste os arquivos CSV das presenças";
DropzonePresenca = new Dropzone("#dropzone-presenca", configDropzone);

configDropzone['dictDefaultMessage'] = "Arraste o arquivo CSV das incrições";
configDropzone['maxFiles'] = 1;
DropzoneAlunos = new Dropzone("#dropzone-alunos", configDropzone);

document.addEventListener('DOMContentLoaded', function() {
    const tooltipTriggerList = document.querySelectorAll('[data-bs-toggle="tooltip"]')
    const tooltipList = [...tooltipTriggerList].map(tooltipTriggerEl => new bootstrap.Tooltip(tooltipTriggerEl))
});

const checkbox = document.getElementById("checkbox")
checkbox.addEventListener("change", () => {
    if (document.body.getAttribute("data-bs-theme") === "light"){
        document.body.setAttribute("data-bs-theme", "dark");
    }else{
        document.body.setAttribute("data-bs-theme", "light");
    }
  
})
