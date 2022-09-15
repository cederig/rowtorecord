import { parseArgs } from 'node:util';
import chalk from 'chalk';
import cliProgress from 'cli-progress';
import XlsxPopulate from 'xlsx-populate';
import yaml from 'js-yaml';
import fs from 'fs';

const bar = new cliProgress.SingleBar({}, cliProgress.Presets.shades_classic);
const log = console.log;

const {
  values: { templateFile, sourceFile, outputFile, mappingFile, verbose },
} = parseArgs({
  options: {
    templateFile: {
      type: "string",
      short: "t",
    },
    sourceFile: {
      type: "string",
      short: "s",
    },
    outputFile: {
      type: "string",
      short: "o",
    },
    mappingFile: {
      type: "string",
      short: "m",
    },
    verbose: {
        type: "boolean",
        short: "v",
      }, 
  },
});

async function converter() {
  
  try {

    if (!templateFile) throw 'Template file not found';
    let modelWorkbook =  await XlsxPopulate.fromFileAsync(templateFile);
    if (!sourceFile) throw 'Source file not found';
    let sourceWorkbook = await XlsxPopulate.fromFileAsync(sourceFile);

    if (!mappingFile) throw 'Mapping file not found';
    const { sheets, modelSheetName } = yaml.load(fs.readFileSync(mappingFile, 'utf8'));
    
    Object.values(sheets).map(sheet => {
        const { sheet: {name, domain, startRow, stopRow, referenceColumn, recordState, mapping } } = sheet;

        if (!modelSheetName) throw 'Missing model sheet name parameter';    
        let mws = modelWorkbook.sheet(modelSheetName);

        if (!name) throw 'Missing source sheet name parameter';    
        let sws = sourceWorkbook.sheet(name);

        const rowCount = stopRow ?? sws._rows.length;
        verbose && bar.start(rowCount-1, 0);

        if (!startRow) throw 'Missing start row parameter';
        if (!referenceColumn) throw 'Missing reference column parameter';

        for(let i = startRow; i < rowCount; i++) {
            verbose && bar.update(i);
            
            let sheetName = sws.row(i).cell(referenceColumn).value();
            if (!sheetName) break;

            let clone = modelWorkbook.cloneSheet(mws, sheetName);
            
            if (recordState) clone.cell("E6").value(recordState);
            if (domain) clone.cell("C9").value(domain);

            Object.values(mapping).map(e => {
                clone.cell(e.target).value(sws.row(i).cell(e.source).value());
            });

            clone.activeCell("A1");
        }
    });
    
    verbose && bar.stop();

    let workbookSheets = modelWorkbook.sheets();
    modelWorkbook.activeSheet(workbookSheets[0]);
    workbookSheets[0].activeCell("A1");
    workbookSheets[1].activeCell("A1");
    modelWorkbook.toFileAsync(outputFile);
    
    verbose && log(chalk.green('File successfully generated: '));
    verbose && log(chalk.green(outputFile));

  } catch (e) {
    verbose && bar.stop();
    console.error(chalk.red('Error: ', e));
  }
}

converter();