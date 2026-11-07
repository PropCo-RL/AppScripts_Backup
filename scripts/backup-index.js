function getProjectFolderName(name) {
  return name.replace(/[^a-zA-Z0-9_-]/g, '_');
}

function extractBackupProjects(rows) {
  const projects = [];
  const projectsById = new Map();
  const seenFolderNames = new Set();

  for (const [index, row] of rows.entries()) {
    const projectName = String(row[0] || '').trim();
    const scriptId = String(row[4] || '').trim();
    const rowNumber = index + 1;

    if (!projectName && !scriptId) continue;
    if (/script\s*id/i.test(scriptId)) continue;
    if (!projectName || !/^[a-zA-Z0-9_-]{20,}$/.test(scriptId)) {
      throw new Error(`Ungültiger Projektname oder Script-ID in Index-Zeile ${rowNumber}.`);
    }

    const folderName = getProjectFolderName(projectName);
    const existingProject = projectsById.get(scriptId);
    if (existingProject) {
      if (existingProject.projectName === projectName) continue;
      throw new Error(`Script-ID ${scriptId} hat mehrere Projektnamen im Backup-Index.`);
    }
    if (seenFolderNames.has(folderName)) {
      throw new Error(`Mehrere Projektnamen ergeben denselben Ordner: ${folderName}.`);
    }

    const project = { projectName, scriptId, folderName };
    projectsById.set(scriptId, project);
    seenFolderNames.add(folderName);
    projects.push(project);
  }

  if (projects.length === 0) {
    throw new Error('Im Tabellenblatt „Übersicht“ wurden keine Projekte mit Script-ID gefunden.');
  }

  return projects;
}

module.exports = { extractBackupProjects, getProjectFolderName };
