package de.intranda.goobi.plugins;

import java.io.File;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.sql.Connection;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.stream.Stream;
import java.util.zip.ZipOutputStream;

import javax.faces.context.ExternalContext;
import javax.faces.context.FacesContext;

import org.apache.commons.configuration.HierarchicalConfiguration;
import org.apache.commons.configuration.XMLConfiguration;
import org.apache.commons.configuration.tree.xpath.XPathExpressionEngine;
import org.apache.commons.dbutils.QueryRunner;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.goobi.beans.Process;
import org.goobi.beans.Processproperty;
import org.goobi.beans.Step;
import org.goobi.production.enums.PluginType;
import org.goobi.production.plugin.interfaces.IWorkflowPlugin;
import org.goobi.vocabulary.Field;
import org.goobi.vocabulary.VocabRecord;

import de.intranda.digiverso.normdataimporter.NormDataImporter;
import de.intranda.digiverso.normdataimporter.model.MarcRecord;
import de.intranda.digiverso.normdataimporter.model.MarcRecord.DatabaseUrl;
import de.sub.goobi.config.ConfigPlugins;
import de.sub.goobi.helper.CloseStepHelper;
import de.sub.goobi.helper.FacesContextHelper;
import de.sub.goobi.helper.Helper;
import de.sub.goobi.helper.StorageProvider;
import de.sub.goobi.helper.enums.StepStatus;
import de.sub.goobi.helper.exceptions.DAOException;
import de.sub.goobi.helper.exceptions.SwapException;
import de.sub.goobi.persistence.managers.MySQLHelper;
import de.sub.goobi.persistence.managers.ProcessManager;
import de.sub.goobi.persistence.managers.ProjectManager;
import de.sub.goobi.persistence.managers.VocabularyManager;
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
    private String closeStepName;
    @Getter
    private boolean exportPossible = false;
    @Getter
    private boolean stepsComplete = false;
    @Getter
    private String projectValidationError = null;
    @Setter
    private String exportFolder;
    @Setter
    private String imageFolder = "media";
    @Getter
    private String projectSizeMessage = null;
    @Getter
    private boolean allowZipDownload = true;

    // used for tests
    @Setter
    private boolean testDatabase;
    @Setter
    private List<Process> testList;

    @Getter
    private boolean includeAllFinishedProcesses = false;

    /**
     * Getter to list all existing active projects
     * 
     * @return List of Strings
     */
    public List<String> getAllProjectNames() {
        if (allProjectNames == null) {
            allProjectNames = ProjectManager.getAllProjectTitles(true);
        }
        return allProjectNames;
    }

    /**
     * Setter to define the project to use
     * 
     * @param selectedProjectName the title of the project to use
     */
    public void setProjectName(String selectedProjectName) {
        if (StringUtils.isBlank(this.projectName) || !this.projectName.equals(selectedProjectName)) {
            this.projectName = selectedProjectName;
            calculateProjectSize();
        }
    }

    /**
     * Select, if all processes of a project are to be included or only the unfinished ones
     * 
     * @param includeAllFinishedProcesses
     */

    public void setIncludeAllFinishedProcesses(boolean includeAllFinishedProcesses) {
        if (this.includeAllFinishedProcesses != includeAllFinishedProcesses) {
            this.includeAllFinishedProcesses = includeAllFinishedProcesses;
            calculateProjectSize();
        }
    }

    private void calculateProjectSize() {
        readConfiguration(projectName);
        projectSizeMessage = null;
        int projectSize = getProcessList().size();
        if (projectSize == 0) {
            stepsComplete = false;
            exportPossible = false;
            projectValidationError = Helper.getTranslation("plugin_workflow_projectexport_emptyProject", projectName);
            Helper.setFehlerMeldung("project", projectValidationError, projectValidationError);
        } else {
            int numberOfTasks = getNumberOfUnfinishedTasks();
            exportPossible = true;
            if (numberOfTasks == 0) {
                stepsComplete = true;
                projectValidationError = null;
                projectSizeMessage = Helper.getTranslation("plugin_workflow_projectexport_projectSize", String.valueOf(projectSize));
            } else {
                // change warning text ?
                stepsComplete = false;
                projectSizeMessage = Helper.getTranslation("plugin_workflow_projectexport_projectSize", String.valueOf(projectSize));
                projectValidationError = Helper.getTranslation("plugin_workflow_projectexport_openSteps", projectName, String.valueOf(numberOfTasks));
                Helper.setFehlerMeldung("project", projectValidationError, projectValidationError);
            }
        }
    }

    /**
     * Find out how many processes in the project are still not in the right status to be interpreted as finished
     * 
     * @return integer value with number of processes
     */
    private int getNumberOfUnfinishedTasks() {
        if (!testDatabase) {
            Connection connection = null;
            StringBuilder query = new StringBuilder();
            query.append("SELECT COUNT(*) ");
            query.append("FROM prozesse WHERE ");
            query.append("istTemplate = false and projekteId = (SELECT projekteID FROM projekte WHERE titel = ?) ");
            query.append("AND prozesseID NOT IN ( ");
            query.append("SELECT prozesseId FROM schritte WHERE titel = ? AND (Bearbeitungsstatus = 3 OR Bearbeitungsstatus = 5)) ");
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

    /**
     * Create a list of all processes of the selected project based on the project title
     * 
     * @return List of processes
     */
    private List<Process> getProcessList() {
        if (!testDatabase) {
            StringBuilder query = new StringBuilder();
            query.append("prozesse.istTemplate = false and projekte.titel = '");
            query.append(projectName);
            query.append("' ");
            query.append("AND prozesseID IN ( ");
            query.append("SELECT prozesseId FROM schritte WHERE titel = '");
            query.append(finishStepName);
            query.append("' ");
            query.append(" AND (Bearbeitungsstatus = 3)) ");
            if (!includeAllFinishedProcesses) {
                query.append("AND prozesseID NOT IN ( ");
                query.append("SELECT prozesseId FROM schritte WHERE titel = '");
                query.append(closeStepName);
                query.append("' AND (Bearbeitungsstatus = 3 OR Bearbeitungsstatus = 5)) ");
            }

            return ProcessManager.getProcesses("prozesse.titel", query.toString(), null);
        } else {
            return testList;
        }
    }

    /**
     * private method to read in all parameters from the configuration file
     * 
     * @param projectName
     */
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
        closeStepName = config.getString("/closeStepName");
        imageFolder = config.getString("/imageFolder", "media");
        allowZipDownload = config.getBoolean("/allowZipDownload", true);
        if (StringUtils.isBlank(exportFolder)) {
            exportFolder = config.getString("/exportDirectory");
        }
    }

    /**
     * Execute the export to write the excel file and the images to the given export folder
     */
    public void prepareExport() {
        // first try to delete previous project exports
        try {
            Path exporttarget = Paths.get(exportFolder, projectName);
            if (Files.exists(exporttarget)) {
                try (Stream<Path> walkStream = Files.walk(exporttarget)) {
                    walkStream.sorted(Comparator.reverseOrder()).map(Path::toFile).forEach(File::delete);
                }
            }
        } catch (IOException e) {
            log.error("Error while deleting previous export results", e);
        }

        List<Process> processesInProject = getProcessList();
        //Properties:
        //    Marginalia  N
        //    Censorship  N
        //    Provenance  Y
        //    Template    Book_workflow_original
        //    TemplateID  54
        //
        //    Book is important
        //    Reason for insignificance
        //
        //    NLI identifier
        //    Reason for missing NLI identifier
        //    OCLC identifier

        // create excel file
        Runnable run = () -> {
            Workbook wb = new SXSSFWorkbook(20);
            Sheet sheet = wb.createSheet("images");

            // create header
            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("File path");
            headerRow.createCell(1).setCellValue("Shots sequence");
            headerRow.createCell(2).setCellValue("Prime Image Flag");
            headerRow.createCell(3).setCellValue("Order");
            headerRow.createCell(4).setCellValue("Identification");
            headerRow.createCell(5).setCellValue("Author lat");
            headerRow.createCell(6).setCellValue("Author in Hebrew");
            headerRow.createCell(7).setCellValue("Other Name Forms");
            headerRow.createCell(8).setCellValue("Litle lat");
            headerRow.createCell(9).setCellValue("Title heb");
            headerRow.createCell(10).setCellValue("NLI number");
            headerRow.createCell(11).setCellValue("OCLC number");
            headerRow.createCell(12).setCellValue("Notes_01");
            headerRow.createCell(13).setCellValue("Normalised Year");
            headerRow.createCell(14).setCellValue("Normalised City");
            headerRow.createCell(15).setCellValue("Reference forms of city.");
            headerRow.createCell(16).setCellValue("Normalised Publisher");
            headerRow.createCell(17).setCellValue("Other name forms for the Publisher");
            headerRow.createCell(18).setCellValue("Notes_02");
            headerRow.createCell(19).setCellValue("Link 1 NLI catalog");
            headerRow.createCell(20).setCellValue("Etichetta 1");
            headerRow.createCell(21).setCellValue("Link 2 website of keeping institution");
            headerRow.createCell(22).setCellValue("Etichetta 2 keeping institution");
            headerRow.createCell(23).setCellValue("Fondo");
            headerRow.createCell(24).setCellValue("Provenance");
            headerRow.createCell(25).setCellValue("Marginalia");
            headerRow.createCell(26).setCellValue("Censorship");
            headerRow.createCell(27).setCellValue("Additional authors in Latin");
            headerRow.createCell(28).setCellValue("Additional authors in Hebrew");
            headerRow.createCell(29).setCellValue("Additional authors references");
            headerRow.createCell(30).setCellValue("Number of copies");
            headerRow.createCell(31).setCellValue("Segnatura");

            int rowCounter = 1;
            boolean error = false;
            processloop: for (Process p : processesInProject) {
                //do this so the metadata is not kept in memory for every process in the list
                Process process = ProcessManager.getProcessById(p.getId());
                // just use this process if the step to check is in valid status
                for (Step step : process.getSchritte()) {
                    if (finishStepName.equals(step.getTitel()) && step.getBearbeitungsstatusEnum() == StepStatus.DEACTIVATED) {
                        continue processloop;
                    }

                }
                log.info("Collect metadata for process {}", process.getTitel());
                // open mets file
                try {
                    Fileformat fileformat = process.readMetadataFile();

                    DigitalDocument digDoc = fileformat.getDigitalDocument();

                    DocStruct logical = digDoc.getLogicalDocStruct();
                    DocStruct physical = digDoc.getPhysicalDocStruct();
                    String representative = "";
                    if (physical.getAllMetadata() != null) {
                        for (Metadata md : physical.getAllMetadata()) {
                            if ("_representative".equals(md.getType().getName())) {
                                representative = md.getValue();
                            }
                        }
                    }
                    // create row for each image
                    List<String> filenames = StorageProvider.getInstance().list(process.getImagesTifDirectory(false));

                    if (!filenames.isEmpty()) {
                        //                if (physical.getAllChildren() != null) {

                        String censorship = "";
                        String marginalia = "";
                        String provenance = "";
                        String oclcIdentifier = "";
                        String notes_01 = "";
                        String titleLat = "";
                        String notes_02 = "";
                        String copies = "";
                        String title = "";
                        String identifier = "";
                        String shelfmark = process.getTitel();
                        String authorLat = "";
                        String authorHeb = "";
                        String authorOther = "";
                        String year = "";
                        String city = "";
                        String cityNormed = "";
                        String cityOther = "";

                        String publisherLat = "";
                        //                    String publisherHeb = "";
                        String publisherOther = "";
                        String nliLink = "";

                        for (Processproperty prop : process.getEigenschaften()) {
                            if ("Censorship".equals(prop.getTitel())) {
                                censorship = prop.getWert();
                            } else if ("Marginalia".equals(prop.getTitel())) {
                                marginalia = prop.getWert();
                            } else if ("Provenance".equals(prop.getTitel())) {
                                provenance = prop.getWert();
                            } else if ("Number of Copies".equals(prop.getTitel())) {
                                copies = prop.getWert();
                            } else if ("NLI_Number".equals(prop.getTitel())) {
                                identifier = prop.getWert();
                            }

                        }

                        StringBuilder additionalAuthorHeb = new StringBuilder();
                        StringBuilder additionalAuthorLat = new StringBuilder();
                        StringBuilder additionalAuthorOther = new StringBuilder();

                        for (Metadata md : logical.getAllMetadata()) {
                            if ("TitleDocMain".equals(md.getType().getName())) {
                                title = md.getValue();
                            } else if ("OtherTitle".equals(md.getType().getName())) {
                                titleLat = md.getValue();
                            } else if ("OclcID".equals(md.getType().getName())) {
                                oclcIdentifier = md.getValue();
                            } else if ("Notes01".equals(md.getType().getName())) {
                                notes_01 = md.getValue();
                            } else if ("Notes02".equals(md.getType().getName())) {
                                notes_02 = md.getValue();
                            } else if ("shelfmarksource".equals(md.getType().getName()) && StringUtils.isNotBlank(md.getValue())) {
                                shelfmark = md.getValue();
                            } else if ("AuthorPreferred".equals(md.getType().getName())) {
                                authorLat = md.getValue();
                            } else if ("AuthorPreferredHeb".equals(md.getType().getName())) {
                                authorHeb = md.getValue();
                            } else if ("AuthorPreferredOther".equals(md.getType().getName())) {
                                authorOther = md.getValue();
                            } else if ("PublicationRun".equals(md.getType().getName())) {
                                year = md.getValue();
                            } else if ("PublicationYear".equals(md.getType().getName())) {
                                year = md.getValue();
                            } else if ("PlaceOfPublicationNormalized".equals(md.getType().getName())) {
                                cityNormed = md.getValue();
                            } else if ("PlaceOfPublication".equals(md.getType().getName())) {
                                city = md.getValue();
                            } else if ("PlaceOfPublicationOther".equals(md.getType().getName())) {
                                cityOther = md.getValue();
                            } else if ("Publisher".equals(md.getType().getName())) {
                                publisherLat = md.getValue();

                                // once we found the publisher name get other writing forms from Vocabulary
                                String vocabRecordUrl = md.getAuthorityValue();
                                if (vocabRecordUrl != null && vocabRecordUrl.length() > 0) {
                                    String vocabID = vocabRecordUrl.substring(vocabRecordUrl.lastIndexOf("/") + 1);
                                    vocabRecordUrl = vocabRecordUrl.substring(0, vocabRecordUrl.lastIndexOf("/"));
                                    String vocabRecordID = vocabRecordUrl.substring(vocabRecordUrl.lastIndexOf("/") + 1);
                                    VocabRecord vr = VocabularyManager.getRecord(Integer.parseInt(vocabRecordID), Integer.parseInt(vocabID));
                                    if (vr != null) {
                                        String url = null;
                                        String value = null;
                                        for (Field f : vr.getFields()) {
                                            if ("Name variants".equals(f.getDefinition().getLabel())) {
                                                publisherOther = f.getValue();
                                            } else if ("Authority URI".equals(f.getDefinition().getLabel())) {
                                                url = f.getValue();
                                            } else if ("Value URI".equals(f.getDefinition().getLabel())) {
                                                value = f.getValue();
                                            }
                                        }

                                        if (StringUtils.isNotBlank(url) && StringUtils.isNotBlank(value) && url.contains("viaf")) {
                                            url = url + value + "/marc21.xml";
                                            MarcRecord recordToImport = null;
                                            try {
                                                recordToImport = NormDataImporter.getSingleMarcRecord(url);
                                            } catch (Exception e) {
                                                log.error(e);
                                            }
                                            if (recordToImport != null) {
                                                List<String> databases = new ArrayList<>();
                                                databases.add("j9u"); // NLI
                                                databases.add("lc"); // LOC
                                                databases.add("bav"); // Vatican
                                                databases.add("gnd"); // GND
                                                databases.add("isni"); // ISNI
                                                DatabaseUrl currentUrl = null;
                                                for (String database : databases) {
                                                    if (currentUrl == null) {
                                                        for (DatabaseUrl dbUrl : recordToImport.getAuthorityDatabaseUrls()) {
                                                            if (dbUrl.getDatabaseCode().equalsIgnoreCase(database)) {
                                                                currentUrl = dbUrl;
                                                            }
                                                        }
                                                    }
                                                }
                                                if (currentUrl == null && !recordToImport.getAuthorityDatabaseUrls().isEmpty()) {
                                                    currentUrl = recordToImport.getAuthorityDatabaseUrls().get(0);
                                                }

                                                if (currentUrl != null) {
                                                    recordToImport = NormDataImporter.getSingleMarcRecord(currentUrl.getMarcRecordUrl());
                                                    if (recordToImport != null) {
                                                        List<String> normalizedVariant =
                                                                recordToImport.getSubFieldValues("100", null, null, "a", "b", "c");
                                                        List<String> otherVariants =
                                                                recordToImport.getSubFieldValues("400", null, null, "a", "b", "c");
                                                        if (normalizedVariant != null) {
                                                            publisherLat = normalizedVariant.get(0);
                                                        }
                                                        if (otherVariants != null) {
                                                            StringBuilder sb = new StringBuilder();
                                                            for (String spelling : otherVariants) {
                                                                if (sb.length() > 0) {
                                                                    sb.append("; ");
                                                                }
                                                                sb.append(spelling);
                                                            }
                                                            if (sb.length() > 0) {
                                                                publisherOther = sb.toString();
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                //                        } else if (md.getType().getName().equals("PublisherHeb")) {
                                //                            publisherHeb = md.getValue();
                            } else if ("NLICatalog".equals(md.getType().getName())) {
                                nliLink = md.getValue();
                            } else if ("AdditionalAuthor".equals(md.getType().getName())) {
                                if (additionalAuthorLat.length() > 0) {
                                    additionalAuthorLat.append("; ");
                                }
                                additionalAuthorLat.append(md.getValue());

                            } else if ("AdditionalAuthorHeb".equals(md.getType().getName())) {
                                if (additionalAuthorHeb.length() > 0) {
                                    additionalAuthorHeb.append("; ");
                                }
                                additionalAuthorHeb.append(md.getValue());

                            } else if ("AdditionalAuthorOther".equals(md.getType().getName())) {
                                if (additionalAuthorOther.length() > 0) {
                                    additionalAuthorOther.append("; ");
                                }
                                additionalAuthorOther.append(md.getValue());
                            }
                        }
                        for (int i = 0; i < filenames.size(); i++) {
                            String imageName = filenames.get(i);

                            String physPageNo = String.valueOf(i + 1);

                            Row imageRow = sheet.createRow(rowCounter);
                            // Field: file path
                            // Comments: A line should be produced for each image taken.
                            // Clarification: Generated by Goobi from the process title/image filename
                            // Example: RM0166F05-0000001/ RM0166F05-0000001_001.jpg
                            imageRow.createCell(0).setCellValue(process.getTitel() + "/" + imageName);
                            // Field: shots sequence
                            // Comments:
                            // Clarification: Generated by Goobi from the image sequence numbering at the end of the image filename
                            // Example: 1
                            imageRow.createCell(1).setCellValue(physPageNo);
                            // Field: Prime Image Flag
                            // Comments: Y/N for the image that should be used as the thumbnail, will be chosen by the cataloger
                            // Clarification: As selected in the workflow by the cataloguer
                            // Example: N
                            imageRow.createCell(2)
                                    .setCellValue(StringUtils.isNotBlank(representative) && representative.equals(physPageNo) ? "Y" : "N");
                            // Field: order
                            // Comments:
                            // Clarification: Generated by Goobi from the process title
                            // Example: RM0166F05-0000001
                            imageRow.createCell(3).setCellValue(process.getTitel());
                            // Field: identification
                            // Comments: Should be the number in the original library. will be loaded based on the excel provided by the institution after inserting the barcodes
                            // Clarification: Imported into Goobi as part of the excel upload of the inventory spreadsheet. This is the shelf mark information provided by the source library (if they use shelf marks) it will not always be present.
                            // Example: CB_FI_015
                            imageRow.createCell(4).setCellValue(process.getTitel());
                            // Field: author lat
                            // Comments: will be taken from VIAF based on the 100 field in the NLI record
                            // Clarification: To be taken from VIAF, Exact Query using the Israel data set on VIAF only the search term is the content of field G below which has been extracted from the NLI ALMA bibliographic record. The version to be used is Either: Italian, Vatican or LOC if Italian or vatican name forms are not present
                            // Example: Aaron Berechiah ben Moses, of Modena, 1549-1639
                            imageRow.createCell(5).setCellValue(authorLat);
                            // Field: Author in Hebrew
                            // Comments: Wil be taken from the 100 field in the NLI record
                            // Clarification: Taken automatically by Goobi from the 100 field in the NLI Alma bibliographic record with the prefix $$HEB (to denote the Hebrew Name)
                            // Example: מודנה, אהרן ברכיה בן משה
                            imageRow.createCell(6).setCellValue(authorHeb);
                            // Field: Other Name Forms
                            // Comments: will be taken from VIAF
                            // Clarification: To be taken from VIAF, Exact Query using the Israel data set on VIAF only the search term is the content of field G above which has been extracted from the NLI ALMA bibliographic record. All other name forms to be copied into this field separated by a semicolon+space "; "
                            // Example: Aaron Berechiah ben Moses von Modena -1639; Aaron Berechja di Modena
                            imageRow.createCell(7).setCellValue(authorOther);
                            // Field: title lat
                            // Comments: will be taken from the OCLC record, or from the manual transliteration.
                            // Clarification: To be taken from the WorldCat transliterated MARC record (field 245) based on the OCLC number entered by the cataloguer (see field L below). If no OCLC number then this will be manually transliterated by the cataloguer
                            // Example: Maʻavar Yaboḳ
                            imageRow.createCell(8).setCellValue(titleLat);
                            // Field: title heb
                            // Comments: should be taken from the 245 feild in the NLI record
                            // Clarification: Taken automatically by Goobi from the 245 field in the NLI Alma bibliographic record
                            // Example: ספר מעבר יבק / שפתי רננות ... עתר ענן הקטרת ... אשר יסד ... כמוהר"ר אהרן ברכיה בכמה"ר משה ממודינה ... בו ביאר איך יתנהג האדם בעה"ז עד עת בוא יום פקודתו ... וחלק אותו לד' חלקים ... שפתי צדק ... שפת אמת ...
                            imageRow.createCell(9).setCellValue(title);
                            // Field: NLI number
                            // Comments: is inserted by the cataloger
                            // Clarification: In most cases this is inserted by the cataloguer in Goobi Workflow after they have found the book on the ALMA system. In some cases this will be inserted by the NLI cataloguer after they have catalogued a new book on ALMA (this is for situations when the cataloguer cannot find the book on ALMA)
                            // Example: 990010919220205000
                            imageRow.createCell(10).setCellValue(identifier);
                            // Field: OCLC number
                            // Comments: is inserted by the cataloger
                            // Clarification: This is inserted by the cataloguer in Goobi workflow if a suitable transliterated record can be found on WorldCat. If not then the book will be manually transliterated by the cataloguer and this field will remain empty
                            // Example: 47085556
                            imageRow.createCell(11).setCellValue(oclcIdentifier);
                            // Field: notes_01
                            // Comments: is taken from the 260 field in OCLC or compiled by the cataloguer if missing
                            // Clarification: This is the imprint field which will be taken from the OCLC record under field 260 (for the majority of the time) or 264 if there is no information in the 260 field
                            // Example: Manṭovah :  Be-vet Yehudah Shemuʼel mi-Prushah u-veno,   [386] 1626.
                            imageRow.createCell(12).setCellValue(notes_01);
                            // Field: Normalised Year
                            // Comments: should be taken from the 008 field in the NLI record
                            // Clarification: To be taken from the NLI ALMA bibliographic record from field 008
                            // Example: 1626
                            imageRow.createCell(13).setCellValue(year);
                            // Field: Normalised City
                            // Comments: should be taken from VIAF (Italian form) based on the 751 NLI record with sub-field "e"="publishing place"
                            // Clarification: To be taken from VIAF, Exact Query using the Israel data set on VIAF only the search term is the content of field 751 (with a sub field "e" which means publishing place) to be extracted from the NLI ALMA bibliographic record. The version to be used is Either: Italian, Vatican or LOC if Italian or vatican name forms are not present
                            // Example: Mantova
                            imageRow.createCell(14).setCellValue(StringUtils.isNotBlank(cityNormed) ? cityNormed : city);
                            // Field: reference forms of city.
                            // Comments: should be taken from VIAF based on the 751 NLI record with sub-field "e"="publishing place"
                            // Clarification: To be taken from VIAF, Exact Query using the Israel data set on VIAF only the search term is the content of field 751 (with a sub field "e" which means publishing place) to be extracted from the NLI ALMA bibliographic record.  All other name forms to be copied into this field separated by a semicolon+space "; "
                            // Example: Mantua (Italy); Mantoue (Italie); מנטובה (איטליה)
                            imageRow.createCell(15).setCellValue(cityOther);
                            // Field: Normalised Publisher
                            // Comments: should be taken from the 7001 or 7102 NLI record with sub-field "e"="publisher"
                            // Clarification: To be taken from the vocabulary manager in Goobi. The cataloguing team has provided approx. 300 publishers, some with VIAF identifiers, to be uploaded to Goobi vocabulary manager. These will therefore need to be manually selected from a drop down list within Goobi Workflow by the cataloguers. As more publishers are added to VIAF the vocabulary can be updated with VIAF identifiers.
                            // Example: Perugia, Yehudah Shemuʼel ben Yehoshuʻa
                            imageRow.createCell(16).setCellValue(publisherLat);
                            // Field: Other name forms for the Publisher
                            // Comments: should be taken from VIAF based on the 751 NLI record with sub-field "e"="publisher
                            // Clarification:
                            // Example:
                            imageRow.createCell(17).setCellValue(publisherOther);
                            // Field: notes_02
                            // Comments: IT IS COMPILED BY THE CATALOGUER ACCORDING TO THE COPY INFORMATION
                            // Clarification: This is an area for the cataloguer to record any notes as needed in Goobi workflow
                            // Example: Missing pages.
                            imageRow.createCell(18).setCellValue(notes_02);
                            // Field: Link 1 NLI catalog
                            // Comments:
                            // Clarification: This is the link to the NLI ALMA catalogue record for the book. Goobi to automatically generate it by combining standard URL prefix with the NLI ALMA number in field K above
                            // Example:
                            imageRow.createCell(19).setCellValue(nliLink);
                            // Field: Etichetta 1
                            // Comments:
                            // Clarification: Standard wording, always the same as in the cell on the right
                            // Example: National Library of Israel record
                            imageRow.createCell(20).setCellValue("National Library of Israel record");
                            // Field: Link 2 website of keeping institution
                            // Comments:
                            // Clarification: This is the website of the holding institution. This is to be inserted by Goobi automatically from the Project record (there will be 1 project per institution)
                            // Example:
                            imageRow.createCell(21).setCellValue(process.getProjekt().getMetsRightsOwnerSite());
                            // Field: Etichetta 2 keeping institution
                            // Comments:
                            // Clarification: This is the name of the holding institution. This is to be inserted by Goobi automatically from the Project record (there will be 1 project per institution)
                            // Example:
                            imageRow.createCell(22).setCellValue(process.getProjekt().getMetsRightsOwner());

                            // Field: Fondo
                            imageRow.createCell(23).setCellValue(process.getProjekt().getMetsRightsSponsor());
                            // Field: provenance
                            // Comments: Y/N, will be chosen by the cataloger or provided by the institution in there excel
                            // Clarification: Imported into Goobi as part of the excel upload of the inventory spreadsheet. This is the provenence information provided by the source library "Y" or "N" it will always be present.
                            // Example: y
                            imageRow.createCell(24).setCellValue(provenance);
                            // Field: Marginalia
                            // Comments: Y/N, will be chosen by the cataloger or provided by the institution in there excel
                            // Clarification: Imported into Goobi as part of the excel upload of the inventory spreadsheet. This is the marginalia information provided by the source library "Y" or "N" it will always be present.
                            // Example: y
                            imageRow.createCell(25).setCellValue(marginalia);
                            // Field: censorship
                            // Comments: Y/N, will be chosen by the cataloger or provided by the institution in there excel
                            // Clarification: Imported into Goobi as part of the excel upload of the inventory spreadsheet. This is the Censorshop information provided by the source library "Y" or "N" it will always be present.
                            // Example: N
                            imageRow.createCell(26).setCellValue(censorship);
                            // Field: additional authors in Latin
                            // Comments: will be taken from VIAF based on the 700 field in the NLI record, can be multiple should be seperated with ";"
                            // Clarification: Taken automatically by Goobi from the 700 field in the NLI Alma bibliographic record with the prefix $$LAT (to denote Latin names) All additional author names to be copied into this field separated by a semicolon+space "; "
                            // Example:

                            imageRow.createCell(27).setCellValue(additionalAuthorLat.toString());
                            // Field: Additional authors in Hebrew
                            // Comments: will be taken from the 700 field in the NLI record, can be multiple should be seperated with ";"
                            // Clarification: Taken automatically by Goobi from the 700 field in the NLI Alma bibliographic record with the prefix $$HEB (to denote Hebrew names) All additional author names to be copied into this field separated by a semicolon+space "; "
                            // Example:

                            imageRow.createCell(28).setCellValue(additionalAuthorHeb.toString());
                            // Field: Additional authors references
                            // Comments: will be taken from VIAF based on the 700 field in the NLI record, can be multiple should be seperated with ";"
                            // Clarification: To be taken from VIAF, Exact Query using the Israel data set on VIAF only the search term is the content of field 700 to be extracted from the NLI ALMA bibliographic record. All other name forms to be copied into this field separated by a semicolon+space "; "
                            // Example:
                            imageRow.createCell(29).setCellValue(additionalAuthorOther.toString());
                            // Field: Number of copies
                            // Comments: calculated by GOOBI, or provided by the institutino in there excel
                            // Clarification: This is to be taken from the excel upload of the inventory spreadsheet
                            // Example: 1
                            imageRow.createCell(30).setCellValue(StringUtils.isBlank(copies) ? "" : copies);

                            imageRow.createCell(31).setCellValue(shelfmark);

                            rowCounter = rowCounter + 1;
                        }
                        if (allowZipDownload) {
                            // export images
                            Path source = Paths.get(process.getConfiguredImageFolder(imageFolder));
                            Path target = Paths.get(exportFolder, projectName, process.getTitel());
                            if (!Files.exists(target)) {
                                Files.createDirectories(target);
                            }
                            StorageProvider.getInstance().copyDirectory(source, target);
                        }
                    }
                } catch (ReadException | PreferencesException | IOException | SwapException | DAOException e) {
                    log.error(e);
                    error = true;
                }
            }
            // save/download excel
            Path destination = Paths.get(exportFolder, projectName);
            if (!StorageProvider.getInstance().isFileExists(destination)) {
                try {
                    StorageProvider.getInstance().createDirectories(destination);
                } catch (IOException e) {
                    log.error(e);
                }
            }
            try {
                Path metadataXlsPath = Paths.get(destination.toString(), "metadata.xlsx");
                log.info("Writing metadata file to {}", metadataXlsPath);
                OutputStream out = Files.newOutputStream(metadataXlsPath);
                wb.write(out);
                out.flush();
                out.close();
                wb.close();
            } catch (IOException e) {
                log.error(e);
                error = true;
            }

            // close step if no error occurred
            if (!error) {
                // close steps in separate thread

                for (Process process : processesInProject) {
                    for (Step step : process.getSchritte()) {
                        if (closeStepName.equals(step.getTitel()) && step.getBearbeitungsstatusEnum() != StepStatus.DEACTIVATED
                                && step.getBearbeitungsstatusEnum() != StepStatus.DONE) {
                            CloseStepHelper.closeStep(step, null);
                            // close step via ticket or goobiscript?
                            break;
                        }
                    }
                }
            }
        };
        Thread createExcelAndCloseThread = new Thread(run);
        createExcelAndCloseThread.start();

        // now zip the entire exported project and allow a download
        if (allowZipDownload) {
            try {
                createExcelAndCloseThread.join();
            } catch (InterruptedException e) {
                log.error(e);
                Helper.setFehlerMeldung("Error exporting project. See application log for details");
                return;
            }
            Helper.setMeldung("plugin_workflow_projectexport_exportFinished");
            try {
                FacesContext facesContext = FacesContextHelper.getCurrentFacesContext();
                ExternalContext ec = facesContext.getExternalContext();
                ec.responseReset();
                ec.setResponseContentType("application/zip");

                ec.setResponseHeader("Content-Disposition", "attachment; filename=" + projectName + ".zip");
                OutputStream responseOutputStream = ec.getResponseOutputStream();
                ZipOutputStream out = new ZipOutputStream(responseOutputStream);

                Path project = Paths.get(exportFolder, projectName);
                ExportThread.zipFolder("", project, out);
                out.flush();
                out.close();

                facesContext.responseComplete();
            } catch (IOException e) {
                log.error(e);
            } finally {
            }
        } else {
            Helper.setMeldung("Export started, this might run a while. Check the export folder for results.");
            ExportThread thread = new ExportThread();
            thread.setExportFolder(exportFolder);
            thread.setImageFolder(imageFolder);
            thread.setProjectName(projectName);
            thread.setProcessesInProject(processesInProject);
            thread.setFinishStepName(finishStepName);
            thread.setWaitforThread(createExcelAndCloseThread);
            thread.start();
        }
    }

    public static void main(String[] args) {
        String vocabRecordUrl = "http://localhost:8080/goobi/api/vocabulary/records/3/14";
        String vocabID = vocabRecordUrl.substring(vocabRecordUrl.lastIndexOf("/") + 1);
        vocabRecordUrl = vocabRecordUrl.substring(0, vocabRecordUrl.lastIndexOf("/"));
        String vocabRecordID = vocabRecordUrl.substring(vocabRecordUrl.lastIndexOf("/") + 1);
        int vid = Integer.parseInt(vocabID);
        int vrid = Integer.parseInt(vocabRecordID);
        System.out.println(vocabID + " - " + vid);
        System.out.println(vocabRecordID + " - " + vrid);
    }
}
