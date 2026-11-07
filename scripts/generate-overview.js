const fs = require('fs');
const path = require('path');

function csvCell(value) {
  const text = String(value ?? '');
  const safeText = /^[=+\-@\t\r]/.test(text) ? `'${text}` : text;
  return `"${safeText.replace(/"/g, '""')}"`;
}

function listScriptFiles(directory) {
  return fs.readdirSync(directory, { withFileTypes: true })
    .flatMap(entry => {
      const filePath = path.join(directory, entry.name);
      if (entry.isDirectory()) return listScriptFiles(filePath);
      return entry.isFile() && /\.(gs|js)$/i.test(entry.name) ? [filePath] : [];
    })
    .sort((a, b) => a < b ? -1 : a > b ? 1 : 0);
}

function generateOverview(projectsDir, outputPath) {
  const rows = [
    ['Projekt', 'Script-ID', 'GS-Datei', 'Repository-Pfad']
  ];

  if (fs.existsSync(projectsDir)) {
    const projects = fs.readdirSync(projectsDir, { withFileTypes: true })
      .filter(entry => entry.isDirectory())
      .sort((a, b) => a.name < b.name ? -1 : a.name > b.name ? 1 : 0);

    for (const project of projects) {
      const projectDir = path.join(projectsDir, project.name);
      let scriptId = '';

      try {
        const claspConfig = JSON.parse(
          fs.readFileSync(path.join(projectDir, '.clasp.json'), 'utf8')
        );
        scriptId = claspConfig.scriptId || '';
      } catch {
        // A missing or invalid .clasp.json does not prevent listing project files.
      }

      for (const filePath of listScriptFiles(projectDir)) {
        const scriptFile = path.relative(projectDir, filePath).split(path.sep).join('/');
        const repositoryPath = `projects/${project.name}/${scriptFile}`;
        rows.push([project.name, scriptId, scriptFile, repositoryPath]);
      }
    }
  }

  const csv = rows.map(row => row.map(csvCell).join(',')).join('\r\n') + '\r\n';
  fs.writeFileSync(outputPath, csv, 'utf8');
  return rows.length - 1;
}

module.exports = generateOverview;
