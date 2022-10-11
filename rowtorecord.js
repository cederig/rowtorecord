"use strict";

const spinner = document.getElementById('spinner');

let store = Object.create(null);

const handleConfiguration = async (event) => {
    try {
        inputReset(event.target);

        const file = event.target.files[0];
        if (!file) throw 'File not found';

        const fileContent = await readFileAsText(file);
        const configuration = parseYaml(fileContent);
        if (!configuration) throw 'Configuration file not found';
        store['configuration'] = configuration;

        const { generatedFileName } = configuration;
        if (generatedFileName) {
            store['generatedFile'] = generatedFileName;

            const { extension, name } = getFileNameExtension(generatedFileName);
            if (name && extension) {
                document.querySelector('#generatedFile').value = name;
                document.getElementById("generatedFile").disabled = true;
                document.getElementById('generatedFileExt').innerHTML = `.${extension}`;
            }
        }

        event.target.classList.add('is-valid');
    } catch (error) {
        event.target.classList.add('is-invalid');
        event.target.nextElementSibling.innerHTML = `Error : ${error}`;
    }
}

const getWorkbook = async (event, store) => {
    try {
        const file = event.target.files[0];
        if (!file) throw new Error('Error: File not found');

        try {
            let data = await XlsxPopulate.fromDataAsync(file);
            if (!data) throw new Error('Error: Data not found');
            event.target.classList.add('is-valid');

            const { extension } = getFileNameExtension(file.name);
            if (!extension) throw new Error('Invalid file extension');
            store(data, extension);

        } catch (error) {
            throw new Error("Error: Something went wrong with this file");
        }

    } catch (error) {
        event.target.classList.add('is-invalid');
        event.target.nextElementSibling.innerHTML = `${error.message}`;
    }
}

document.querySelector('#templateFile').addEventListener('change', event => {
    getWorkbook(event, (data, extension) => {
        store['templateWorkbook'] = data;
        store['extension'] = extension;
        document.getElementById('generatedFileExt').innerHTML = `.${extension}`;
    });
});

document.querySelector('#sourceFile').addEventListener('change', event => {
    getWorkbook(event, (data) => {
        store['sourceWorkbook'] = data;
    });
});

document.querySelector('#configurationFile').addEventListener('change', handleConfiguration);

document.querySelector('#generatedFile').addEventListener('change', event => {
    store['generatedFile'] = store['extension'] ? `${event.target.value}.${store['extension']}` : event.target.value;
});

async function converter() {
    try {
        if (!store['templateWorkbook']) {
            document.querySelector('#templateFile').classList.add('is-invalid');
            throw 'Template file not found';
        }
        if (!store['sourceWorkbook']) {
            document.querySelector('#sourceFile').classList.add('is-invalid');
            throw 'Source file not found';
        }
        if (!store['configuration']) {
            document.querySelector('#configurationFile').classList.add('is-invalid');
            throw 'Mapping file not found';
        }
        if (!store['generatedFile']) {
            document.querySelector('#generatedFile').classList.add('is-invalid');
            throw 'Generated file name not found';
        }

        const { sheets, modelSheetName } = store['configuration'];

        spinnerOn(spinner);

        Object.values(sheets).map(sheet => {
            try {
                const { sheet: { name, startRow, stopRow, referenceColumn, mapping } } = sheet;

                if (!modelSheetName) throw new SyntaxError('Missing model sheet name parameter');
                let mws = store['templateWorkbook'].sheet(modelSheetName);
                if (!mws) throw new SyntaxError(`Model sheet name '${modelSheetName}' not found`);

                if (!name) throw new SyntaxError('Missing source sheet name parameter');
                let sws = store['sourceWorkbook'].sheet(name);
                if (!sws) throw new SyntaxError(`Source sheet name '${name}' not found`);

                const rowCount = stopRow ?? sws._rows.length;

                if (!startRow) throw 'Missing start row parameter';
                if (!referenceColumn) throw 'Missing reference column parameter';

                for (let i = startRow; i < rowCount; i++) {

                    let sheetName = sws.row(i).cell(referenceColumn).value();
                    if (!sheetName) break;

                    let clone = store['templateWorkbook'].cloneSheet(mws, sheetName);

                    Object.values(mapping).map(e => {
                        clone.cell(e.target).value(sws.row(i).cell(e.source).value());
                    });

                    clone.activeCell("A1");
                }
            } catch (error) {
                document.getElementById("err").innerHTML = error;
                throw error;
            }
        });

        // let workbookSheets = modelWorkbook.sheets();
        // modelWorkbook.activeSheet(workbookSheets[0]);
        // workbookSheets[0].activeCell("A1");
        // workbookSheets[1].activeCell("A1");

        generate(store['templateWorkbook']).then((blob) => {
            var url = window.URL.createObjectURL(blob);
            var a = document.createElement("a");
            document.body.appendChild(a);
            a.href = url;
            a.download = store['generatedFile'];
            a.click();
            window.URL.revokeObjectURL(url);
            document.body.removeChild(a);

            resetForm();
            // console.log(XlsxPopulate.MIME_TYPE);
            // location.href = "data:" + XlsxPopulate.MIME_TYPE + ";base64," + base64;
        });
    } catch (error) {
        document.getElementById("err").innerHTML = error;
    }
}

const generate = workbook => {
    return workbook.outputAsync();
}

const parseYaml = contents => {
    return jsyaml.load(contents);
}

const readFileAsText = inputFile => {
    const reader = new FileReader();

    return new Promise((resolve, reject) => {
        reader.onerror = () => {
            reader.abort();
            reject(new DOMException("Problem parsing input file."));
        };

        reader.onload = () => {
            resolve(reader.result);
        };
        reader.readAsText(inputFile);
    });
};

const spinnerOn = element => {
    element.removeAttribute("hidden");
    element.parentNode.setAttribute("disabled", "");
}

const spinnerOff = element => {
    element.setAttribute("hidden", "");
    element.parentNode.removeAttribute("disabled");
}



const inputReset = value => {
    value.classList.remove('is-invalid');
    value.classList.remove('is-valid');
}

const resetForm = () => {
    document.getElementById("err").innerHTML = '';
    document.getElementById("generatedFile").disabled = false;
    document.getElementById('generatedFileExt').innerHTML = ".xlsx";

    let t = document.getElementById('templateFile');
    let s = document.getElementById('sourceFile');
    let v = document.getElementById('configurationFile');
    let u = document.getElementById('generatedFile');
    t.value = t.defaultValue;
    s.value = s.defaultValue;
    u.value = u.defaultValue;
    v.value = v.defaultValue;

    inputReset(t);
    inputReset(s);
    inputReset(u);
    inputReset(v);
    store = {};
    spinnerOff(spinner);

}

const getFileNameExtension = (filename) => {
    let last_dot = filename.lastIndexOf('.');
    if (last_dot == -1) throw new SyntaxError('Missing extension file');
    let ext = filename.slice(last_dot + 1);
    let name = filename.slice(0, last_dot);
    return { filename: filename, extension: ext, name: name };
}