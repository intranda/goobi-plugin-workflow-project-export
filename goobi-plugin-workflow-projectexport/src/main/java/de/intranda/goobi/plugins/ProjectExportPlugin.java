package de.intranda.goobi.plugins;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.sql.Connection;
import java.sql.SQLException;
import java.util.List;

import org.apache.commons.configuration.HierarchicalConfiguration;
import org.apache.commons.configuration.XMLConfiguration;
import org.apache.commons.configuration.tree.xpath.XPathExpressionEngine;
import org.apache.commons.dbutils.QueryRunner;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.goobi.beans.Process;
import org.goobi.beans.Processproperty;
import org.goobi.beans.Step;
import org.goobi.production.enums.PluginType;
import org.goobi.production.plugin.interfaces.IWorkflowPlugin;

import de.sub.goobi.config.ConfigPlugins;
import de.sub.goobi.helper.CloseStepHelper;
import de.sub.goobi.helper.Helper;
import de.sub.goobi.helper.StorageProvider;
import de.sub.goobi.helper.enums.StepStatus;
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
    private String closeStepName;

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
        closeStepName = config.getString("/closeStepName");
    }

    public void prepareExport() {

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
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("images");

        // create header
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("file path");
        headerRow.createCell(1).setCellValue("shots sequence");
        headerRow.createCell(2).setCellValue("Prime Image Flag");
        headerRow.createCell(3).setCellValue("order");
        headerRow.createCell(4).setCellValue("identification");
        headerRow.createCell(5).setCellValue("author lat");
        headerRow.createCell(6).setCellValue("Author in Hebrew");
        headerRow.createCell(7).setCellValue("Other Name Forms");
        headerRow.createCell(8).setCellValue("title lat");
        headerRow.createCell(9).setCellValue("title heb");
        headerRow.createCell(10).setCellValue("NLI number");
        headerRow.createCell(11).setCellValue("OCLC number");
        headerRow.createCell(12).setCellValue("notes_01");
        headerRow.createCell(13).setCellValue("Normalised Year");
        headerRow.createCell(14).setCellValue("Normalised City");
        headerRow.createCell(15).setCellValue("reference forms of city.");
        headerRow.createCell(16).setCellValue("Normalised Publisher");
        headerRow.createCell(17).setCellValue("Other name forms for the Publisher");
        headerRow.createCell(18).setCellValue("notes_02");
        headerRow.createCell(19).setCellValue("Link 1 NLI catalog");
        headerRow.createCell(20).setCellValue("Etichetta 1");
        headerRow.createCell(21).setCellValue("Link 2 website of keeping institution");
        headerRow.createCell(22).setCellValue("Etichetta 2 keeping institution");
        headerRow.createCell(23).setCellValue("provenance");
        headerRow.createCell(24).setCellValue("Marginalia");
        headerRow.createCell(25).setCellValue("censorship");
        headerRow.createCell(26).setCellValue("additional authors in Latin");
        headerRow.createCell(27).setCellValue("Additional authors in Hebrew");
        headerRow.createCell(28).setCellValue("Additional authors references");
        headerRow.createCell(29).setCellValue("Number of copies");

        int rowCounter = 1;
        boolean error = false;
        for (Process process : processesInProject) {

            // open mets file
            try {
                Fileformat fileformat = process.readMetadataFile();

                DigitalDocument digDoc = fileformat.getDigitalDocument();

                DocStruct logical = digDoc.getLogicalDocStruct();
                DocStruct physical = digDoc.getPhysicalDocStruct();
                String representative = "";
                if (physical.getAllMetadata() != null) {
                    for (Metadata md : physical.getAllMetadata()) {
                        if (md.getType().getName().equals("_representative")) {
                            representative = md.getValue();
                        }
                    }
                }
                // create row for each image
                if (physical.getAllChildren() != null) {

                    String censorship = "";
                    String marginalia = "";
                    String provenance = "";
                    String oclcIdentifier = "";
                    String notes_01 = ""; // TODO get it from correct field
                    String titleLat = ""; // TODO get it from correct field
                    String notes_02 = ""; // TODO get it from correct field
                    String copies = "";
                    String title = "";
                    String identifier = "";
                    String shelfmark = "";
                    String authorLat = "";
                    String authorHeb = "";
                    String authorOther = "";
                    String year = "";
                    String city = "";
                    String cityNormed = "";
                    String cityOther = "";

                    String publisherLat = "";
                    String publisherHeb = "";
                    String publisherOther = "";
                    String nliLink = "";

                    for (Processproperty prop : process.getEigenschaften()) {
                        if (prop.getTitel().equals("Censorship")) {
                            censorship = prop.getWert();
                        } else if (prop.getTitel().equals("Marginalia")) {
                            marginalia = prop.getWert();
                        } else if (prop.getTitel().equals("Provenance")) {
                            provenance = prop.getWert();
                        } else if (prop.getTitel().equals("OCLC identifier")) {
                            oclcIdentifier = prop.getWert();
                        } else if (prop.getTitel().equals("Number of Copies")) {
                            copies = prop.getWert();
                            //                        } else if (prop.getTitel().equals("notes_01")) {
                            //                            notes01 = prop.getWert();
                        }

                    }

                    StringBuilder additionalAuthorHeb = new StringBuilder();
                    StringBuilder additionalAuthorLat = new StringBuilder();
                    StringBuilder additionalAuthorOther = new StringBuilder();

                    for (Metadata md : logical.getAllMetadata()) {
                        if (md.getType().getName().equals("TitleDocMain")) {
                            title = md.getValue();
                        } else if (md.getType().getName().equals("AlmaID")) {
                            identifier = md.getValue();
                        } else if (md.getType().getName().equals("shelfmarksource")) {
                            shelfmark = md.getValue();
                        } else if (md.getType().getName().equals("AuthorPreferred")) {
                            authorLat = md.getValue();
                        } else if (md.getType().getName().equals("AuthorPreferredHeb")) {
                            authorHeb = md.getValue();
                        } else if (md.getType().getName().equals("AuthorPreferredOther")) {
                            authorOther = md.getValue();
                        } else if (md.getType().getName().equals("PublicationRun")) {
                            year = md.getValue();
                        } else if (md.getType().getName().equals("PublicationYear")) {
                            year = md.getValue();
                        } else if (md.getType().getName().equals("PlaceOfPublicationNormalized")) {
                            cityNormed = md.getValue();
                        } else if (md.getType().getName().equals("PlaceOfPublication")) {
                            city = md.getValue();
                        } else if (md.getType().getName().equals("PlaceOfPublicationOther")) {
                            cityOther = md.getValue();
                        } else if (md.getType().getName().equals("Publisher")) {
                            publisherLat = md.getValue();
                        } else if (md.getType().getName().equals("PublisherHeb")) {
                            publisherHeb = md.getValue();
                        } else if (md.getType().getName().equals("PublisherOther")) {
                            publisherOther = md.getValue();
                        } else if (md.getType().getName().equals("NLICatalog")) {
                            nliLink = md.getValue();
                        } else if (md.getType().getName().equals("AdditionalAuthor")) {
                            if (additionalAuthorLat.length() > 0) {
                                additionalAuthorLat.append("; ");
                            }
                            additionalAuthorLat.append(md.getValue());

                        } else if (md.getType().getName().equals("AdditionalAuthorHeb")) {
                            if (additionalAuthorHeb.length() > 0) {
                                additionalAuthorHeb.append("; ");
                            }
                            additionalAuthorHeb.append(md.getValue());

                        } else if (md.getType().getName().equals("AdditionalAuthorOther")) {
                            if (additionalAuthorOther.length() > 0) {
                                additionalAuthorOther.append("; ");
                            }
                            additionalAuthorOther.append(md.getValue());
                        }
                    }
                    for (DocStruct page : physical.getAllChildren()) {

                        String physPageNo = "";
                        //                        String logPageNo = "";
                        for (Metadata md : page.getAllMetadata()) {
                            if (md.getType().getName().equals("physPageNumber")) {
                                physPageNo = md.getValue();
                                //                            } else if (md.getType().getName().equals("logicalPageNumber")) {
                                //                                logPageNo = md.getValue();
                            }
                        }
                        Row imageRow = sheet.createRow(rowCounter);
                        // Field: file path
                        // Comments: A line should be produced for each image taken.
                        // Clarification: Generated by Goobi from the process title/image filename
                        // Example: RM0166F05-0000001/ RM0166F05-0000001_001.jpg
                        imageRow.createCell(0).setCellValue(process.getTitel() + "/" + page.getImageName());
                        // Field: shots sequence
                        // Comments:
                        // Clarification: Generated by Goobi from the image sequence numbering at the end of the image filename
                        // Example: 1
                        imageRow.createCell(1).setCellValue(physPageNo);
                        // Field: Prime Image Flag
                        // Comments: Y/N for the image that should be used as the thumbnail, will be chosen by the cataloger
                        // Clarification: As selected in the workflow by the cataloguer
                        // Example: N
                        imageRow.createCell(2).setCellValue(StringUtils.isNotBlank(representative) && representative.equals(physPageNo) ? "Y" : "N");
                        // Field: order
                        // Comments:
                        // Clarification: Generated by Goobi from the process title
                        // Example: RM0166F05-0000001
                        imageRow.createCell(3).setCellValue(process.getTitel());
                        // Field: identification
                        // Comments: Should be the number in the original library. will be loaded based on the excel provided by the institution after inserting the barcodes
                        // Clarification: Imported into Goobi as part of the excel upload of the inventory spreadsheet. This is the shelf mark information provided by the source library (if they use shelf marks) it will not always be present.
                        // Example: CB_FI_015
                        imageRow.createCell(4).setCellValue(shelfmark);
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
                        // Field: provenance
                        // Comments: Y/N, will be chosen by the cataloger or provided by the institution in there excel
                        // Clarification: Imported into Goobi as part of the excel upload of the inventory spreadsheet. This is the provenence information provided by the source library "Y" or "N" it will always be present.
                        // Example: y
                        imageRow.createCell(23).setCellValue(provenance);
                        // Field: Marginalia
                        // Comments: Y/N, will be chosen by the cataloger or provided by the institution in there excel
                        // Clarification: Imported into Goobi as part of the excel upload of the inventory spreadsheet. This is the marginalia information provided by the source library "Y" or "N" it will always be present.
                        // Example: y
                        imageRow.createCell(24).setCellValue(marginalia);
                        // Field: censorship
                        // Comments: Y/N, will be chosen by the cataloger or provided by the institution in there excel
                        // Clarification: Imported into Goobi as part of the excel upload of the inventory spreadsheet. This is the Censorshop information provided by the source library "Y" or "N" it will always be present.
                        // Example: N
                        imageRow.createCell(25).setCellValue(censorship);
                        // Field: additional authors in Latin
                        // Comments: will be taken from VIAF based on the 700 field in the NLI record, can be multiple should be seperated with ";"
                        // Clarification: Taken automatically by Goobi from the 700 field in the NLI Alma bibliographic record with the prefix $$LAT (to denote Latin names) All additional author names to be copied into this field separated by a semicolon+space "; "
                        // Example:

                        imageRow.createCell(26).setCellValue(additionalAuthorLat.toString());
                        // Field: Additional authors in Hebrew
                        // Comments: will be taken from the 700 field in the NLI record, can be multiple should be seperated with ";"
                        // Clarification: Taken automatically by Goobi from the 700 field in the NLI Alma bibliographic record with the prefix $$HEB (to denote Hebrew names) All additional author names to be copied into this field separated by a semicolon+space "; "
                        // Example:

                        imageRow.createCell(27).setCellValue(additionalAuthorHeb.toString());
                        // Field: Additional authors references
                        // Comments: will be taken from VIAF based on the 700 field in the NLI record, can be multiple should be seperated with ";"
                        // Clarification: To be taken from VIAF, Exact Query using the Israel data set on VIAF only the search term is the content of field 700 to be extracted from the NLI ALMA bibliographic record. All other name forms to be copied into this field separated by a semicolon+space "; "
                        // Example:
                        imageRow.createCell(28).setCellValue(additionalAuthorOther.toString());
                        // Field: Number of copies
                        // Comments: calculated by GOOBI, or provided by the institutino in there excel
                        // Clarification: This is to be taken from the excel upload of the inventory spreadsheet
                        // Example: 1
                        imageRow.createCell(29).setCellValue(StringUtils.isBlank(copies) ? "1" : copies);

                        rowCounter = rowCounter + 1;
                    }
                    // export images
                    Path source = Paths.get(process.getImagesTifDirectory(false));
                    Path target = Paths.get(exportFolder, identifier);
                    if (!Files.exists(target)) {
                        Files.createDirectories(target);
                    }
                    StorageProvider.getInstance().copyDirectory(source, target);
                }
            } catch (ReadException | PreferencesException | WriteException | IOException | InterruptedException | SwapException | DAOException e) {
                log.error(e);
                error = true;
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
            log.error(e);
            error = true;
        }
        // close step if no error occurred
        if (!error) {
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
    }

}
