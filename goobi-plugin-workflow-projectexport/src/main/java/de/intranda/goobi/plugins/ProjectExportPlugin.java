package de.intranda.goobi.plugins;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.configuration.HierarchicalConfiguration;
import org.apache.commons.configuration.XMLConfiguration;
import org.apache.commons.configuration.tree.xpath.XPathExpressionEngine;
import org.apache.commons.dbutils.QueryRunner;
import org.apache.commons.dbutils.ResultSetHandler;
import org.apache.commons.io.FileUtils;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.goobi.beans.Process;
import org.goobi.production.enums.PluginType;
import org.goobi.production.plugin.interfaces.IWorkflowPlugin;

import de.sub.goobi.config.ConfigPlugins;
import de.sub.goobi.helper.Helper;
import de.sub.goobi.helper.exceptions.DAOException;
import de.sub.goobi.helper.exceptions.SwapException;
import de.sub.goobi.persistence.managers.MySQLHelper;
import de.sub.goobi.persistence.managers.ProcessManager;
import de.sub.goobi.persistence.managers.ProjectManager;
import lombok.Getter;
import lombok.Setter;
import lombok.extern.log4j.Log4j2;
import net.xeoh.plugins.base.annotations.PluginImplementation;
import ugh.dl.DigitalDocument;
import ugh.dl.DocStruct;
import ugh.dl.Fileformat;
import ugh.dl.Metadata;
import ugh.exceptions.PreferencesException;
import ugh.exceptions.ReadException;
import ugh.exceptions.WriteException;

@PluginImplementation
@Log4j2
public class ProjectExportPlugin implements IWorkflowPlugin {

    //Development of a Goobi workflow plugin to manage a collected export of all processes of a project:
    //
    //    Implement a workflow plugin with user interface to see all projects and to trigger an export from there.
    //    Check if all processes of a project have reached a certain step in the workflow to allow/forbid this export to happen
    //    Automatically prepare a zip file with the required compressed images as JPEG with 150 dpi and a maximum file size of 1 MB
    //    Generation of Excel files per project including several metadata per image together with the image name
    //    Allow the download of the prepared zip file per project that includes all results for the ingest into TECA
    //    Create and publish a documentation for this plugin
    //

    @Getter
    private String title = "intranda_workflow_projectexport";

    @Getter
    private PluginType type = PluginType.Workflow;

    @Getter
    private String gui = "/uii/plugin_workflow_projectexport.xhtml";

    @Getter
    private String projectName;

    @Setter
    private List<String> allProjectNames = null;

    @Getter
    private String finishStepName;
    @Getter
    private boolean exportAllowed = false;

    @Getter
    private String projectValidationError = null;

    @Setter
    private String exportFolder;

    // used for tests
    @Setter
    private boolean testDatabase;
    @Setter
    private List<Process> testList;

    public List<String> getAllProjectNames() {
        if (allProjectNames == null) {
            allProjectNames = ProjectManager.getAllProjectTitles(true);
        }
        return allProjectNames;
    }

    public void setProjectName(String selectedProjectName) {
        if (StringUtils.isBlank(this.projectName) || !this.projectName.equals(selectedProjectName)) {
            this.projectName = selectedProjectName;
            readConfiguration(projectName);

            int numberOfTasks = getNumberOfUnfinishedTasks();
            if (numberOfTasks == 0) {
                exportAllowed = true;
                projectValidationError = null;
            } else {
                exportAllowed = false;
                projectValidationError = "The project '" + projectName + "' has " + numberOfTasks + " unfinished processes";
                Helper.setFehlerMeldung("project", projectValidationError, projectValidationError);
            }
        }
    }

    private int getNumberOfUnfinishedTasks() {
        if (!testDatabase) {
            Connection connection = null;
            StringBuilder query = new StringBuilder();
            query.append("SELECT COUNT(*) ");
            query.append("FROM prozesse WHERE ");
            query.append("projekteId = (SELECT projekteID FROM projekte WHERE titel = ?) ");
            query.append("AND prozesseID NOT IN ( ");
            query.append("SELECT prozesseId FROM schritte WHERE titel = ? AND Bearbeitungsstatus = 3) ");
            try {
                connection = MySQLHelper.getInstance().getConnection();
                int currentValue =
                        new QueryRunner().query(connection, query.toString(), MySQLHelper.resultSetToIntegerHandler, projectName, finishStepName);
                return currentValue;
            } catch (SQLException e) {
                log.error(e);
            } finally {
                if (connection != null) {
                    try {
                        MySQLHelper.closeConnection(connection);
                    } catch (SQLException e) {
                        log.error(e);
                    }
                }
            }
        }
        return 0;
    }

    private List<Process> getProcessList() {
        if (!testDatabase) {
            return ProcessManager.getProcesses("prozesse.titel", "projekte.titel = '" + projectName + "' ");
        } else {
            return testList;
        }
    }

    private void readConfiguration(String projectName) {

        HierarchicalConfiguration config = null;
        XMLConfiguration xmlConfig = ConfigPlugins.getPluginConfig(title);
        xmlConfig.setExpressionEngine(new XPathExpressionEngine());

        // order of configuration is:
        //        1.) project name and step name matches
        //        2.) step name matches and project is *
        //        3.) project name matches and step name is *
        //        4.) project name and step name are *
        try {
            config = xmlConfig.configurationAt("//config[./project = '" + projectName + "']");
        } catch (IllegalArgumentException e) {
            try {
                config = xmlConfig.configurationAt("//config[./project = '*']");
            } catch (IllegalArgumentException e1) {

            }
        }
        finishStepName = config.getString("/finishedStepName");
    }

    public void prepareExport() {

        List<Process> processesInProject = getProcessList();

        // create excel file
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("images");

        // create header
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("Image");
        headerRow.createCell(1).setCellValue("Title");
        headerRow.createCell(2).setCellValue("Identifier");
        headerRow.createCell(3).setCellValue("physical page number");
        headerRow.createCell(4).setCellValue("logical page number");

        int rowCounter = 1;

        for (Process process : processesInProject) {

            // open mets file
            try {
                Fileformat fileformat = process.readMetadataFile();

                DigitalDocument digDoc = fileformat.getDigitalDocument();

                DocStruct anchor = null;
                DocStruct logical = digDoc.getLogicalDocStruct();
                if (logical.getType().isAnchor()) {
                    anchor = logical;
                    logical = logical.getAllChildren().get(0);
                }
                DocStruct physical = digDoc.getPhysicalDocStruct();

                // create row for each image
                if (physical.getAllChildren() != null) {
                    String title = "";
                    String identifier = "";
                    for (Metadata md : logical.getAllMetadata()) {
                        if (md.getType().getName().equals("TitleDocMain")) {
                            title = md.getValue();
                        } else if (md.getType().getName().equals("CatalogIDDigital")) {
                            identifier = md.getValue();
                        }
                    }
                    for (DocStruct page : physical.getAllChildren()) {

                        String physPageNo = "";
                        String logPageNo = "";
                        for (Metadata md : page.getAllMetadata()) {
                            if (md.getType().getName().equals("physPageNumber")) {
                                physPageNo = md.getValue();
                            } else if (md.getType().getName().equals("logicalPageNumber")) {
                                logPageNo = md.getValue();
                            }
                        }
                        Row imageRow = sheet.createRow(rowCounter);
                        imageRow.createCell(0).setCellValue(exportFolder + identifier + "/" + page.getImageName());
                        imageRow.createCell(1).setCellValue(title);
                        imageRow.createCell(2).setCellValue(identifier);
                        imageRow.createCell(3).setCellValue(physPageNo);
                        imageRow.createCell(4).setCellValue(logPageNo);
                        rowCounter = rowCounter + 1;
                    }
                    // export images
                    // TODO
                    Path source = Paths.get(process.getImagesTifDirectory(false));
                    Path target = Paths.get(exportFolder, identifier);
                    if (!Files.exists(target)) {
                        Files.createDirectories(target);
                    }
                    FileUtils.copyDirectory(source.toFile(), target.toFile());
                    //                    StorageProvider.getInstance().copyDirectory(Paths.get(process.getImagesTifDirectory(false)), Paths.get(exportFolder + identifier));
                }

            } catch (ReadException | PreferencesException | WriteException | IOException | InterruptedException | SwapException | DAOException e) {
                // TODO Auto-generated catch block
                e.printStackTrace();
            }

        }
        // save/download excel
        try {
            OutputStream out = Files.newOutputStream(Paths.get(exportFolder, "metadata.xlsx"));
            wb.write(out);
            out.flush();
            out.close();
            wb.close();
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
        for (Process process : processesInProject) {

            // TODO close step xy via goobiscript?

        }
    }

    public static ResultSetHandler<Map<Integer, Map<String, String>>> resultSetToProjectMetadata =
            new ResultSetHandler<Map<Integer, Map<String, String>>>() {
        @Override
        public Map<Integer, Map<String, String>> handle(ResultSet rs) throws SQLException {
            Map<Integer, Map<String, String>> projectMetadata = new HashMap<>();
            try {
                while (rs.next()) {
                    Integer id = rs.getInt("processid");
                    String name = rs.getString("name");
                    String value = rs.getString("value");

                    Map<String, String> map = projectMetadata.get(id);
                    if (map == null) {
                        map = new HashMap<>();
                        projectMetadata.put(id, map);
                    }

                }
            } finally {
                if (rs != null) {
                    rs.close();
                }
            }
            return projectMetadata;
        }
    };

}
