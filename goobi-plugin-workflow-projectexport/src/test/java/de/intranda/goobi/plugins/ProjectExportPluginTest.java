package de.intranda.goobi.plugins;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertTrue;

import java.io.File;
import java.io.IOException;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import org.apache.commons.configuration.XMLConfiguration;
import org.apache.commons.configuration.reloading.FileChangedReloadingStrategy;
import org.easymock.EasyMock;
import org.easymock.IAnswer;
import org.goobi.beans.Process;
import org.goobi.beans.Processproperty;
import org.goobi.beans.Project;
import org.goobi.beans.Step;
import org.goobi.beans.User;
import org.junit.Before;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;
import org.junit.runner.RunWith;
import org.powermock.api.easymock.PowerMock;
import org.powermock.core.classloader.annotations.PowerMockIgnore;
import org.powermock.core.classloader.annotations.PrepareForTest;
import org.powermock.modules.junit4.PowerMockRunner;

import de.sub.goobi.config.ConfigPlugins;
import de.sub.goobi.config.ConfigurationHelper;
import de.sub.goobi.helper.CloseStepHelper;
import de.sub.goobi.helper.VariableReplacer;
import de.sub.goobi.helper.enums.PropertyType;
import de.sub.goobi.helper.enums.StepStatus;
import de.sub.goobi.metadaten.MetadatenHelper;
import de.sub.goobi.persistence.managers.ProcessManager;
import de.sub.goobi.persistence.managers.StepManager;
import ugh.dl.Fileformat;
import ugh.dl.Prefs;
import ugh.fileformats.mets.MetsMods;

@RunWith(PowerMockRunner.class)
@PrepareForTest({ MetadatenHelper.class, VariableReplacer.class, ConfigPlugins.class, StepManager.class, ConfigurationHelper.class,
    ProcessManager.class, CloseStepHelper.class })
@PowerMockIgnore({ "javax.management.*" })
public class ProjectExportPluginTest {

    @Rule
    public TemporaryFolder folder = new TemporaryFolder();
    private File processDirectory;
    private File metadataDirectory;
    private Process process;

    private String resourcesFolder;

    @Before
    public void setUp() throws Exception {

        resourcesFolder = "src/test/resources/"; // for junit tests in eclipse

        if (!Files.exists(Paths.get(resourcesFolder))) {
            resourcesFolder = "target/test-classes/"; // to run mvn test from cli or in jenkins
        }

        metadataDirectory = folder.newFolder("metadata");

        processDirectory = new File(metadataDirectory + File.separator + "1");
        processDirectory.mkdirs();
        String metadataDirectoryName = metadataDirectory.getAbsolutePath() + File.separator;

        // copy meta.xml
        Path metaSource = Paths.get(resourcesFolder + "meta.xml");
        Path metaTarget = Paths.get(processDirectory.getAbsolutePath(), "meta.xml");
        Files.copy(metaSource, metaTarget);

        // copy ruleset

        XMLConfiguration config = getConfig();
        PowerMock.mockStatic(ConfigPlugins.class);
        EasyMock.expect(ConfigPlugins.getPluginConfig("intranda_workflow_projectexport")).andReturn(config).anyTimes();

        PowerMock.mockStatic(ConfigurationHelper.class);
        ConfigurationHelper configurationHelper = EasyMock.createMock(ConfigurationHelper.class);
        EasyMock.expect(ConfigurationHelper.getInstance()).andReturn(configurationHelper).anyTimes();
        EasyMock.expect(configurationHelper.getMetsEditorLockingTime()).andReturn(1800000l).anyTimes();
        EasyMock.expect(configurationHelper.isAllowWhitespacesInFolder()).andReturn(false).anyTimes();
        EasyMock.expect(configurationHelper.useS3()).andReturn(false).anyTimes();
        EasyMock.expect(configurationHelper.isUseMasterDirectory()).andReturn(true).anyTimes();
        EasyMock.expect(configurationHelper.isCreateMasterDirectory()).andReturn(false).anyTimes();
        EasyMock.expect(configurationHelper.isCreateSourceFolder()).andReturn(false).anyTimes();
        EasyMock.expect(configurationHelper.getProcessImagesMainDirectoryName()).andReturn("RM0166F05-0000001_media").anyTimes();
        EasyMock.expect(configurationHelper.getFolderForInternalProcesslogFiles()).andReturn("intern").anyTimes();
        EasyMock.expect(configurationHelper.getMetadataFolder()).andReturn(metadataDirectoryName).anyTimes();

        EasyMock.expect(configurationHelper.getScriptCreateDirMeta()).andReturn("").anyTimes();
        EasyMock.replay(configurationHelper);
        PowerMock.mockStatic(VariableReplacer.class);
        EasyMock.expect(VariableReplacer.simpleReplace(EasyMock.anyString(), EasyMock.anyObject())).andReturn("RM0166F05-0000001_media").anyTimes();
        PowerMock.replay(VariableReplacer.class);

        PowerMock.mockStatic(CloseStepHelper.class);
        EasyMock.expect(CloseStepHelper.closeStep(EasyMock.anyObject(), EasyMock.anyObject())).andAnswer(new IAnswer<Boolean>() {

            @Override
            public Boolean answer() throws Throwable {
                for (Step step : process.getSchritte()) {
                    if (step.getTitel().equals("test step to close")) {
                        step.setBearbeitungsstatusEnum(StepStatus.DONE);
                    } else if (step.getTitel().equals("locked step that should open")) {
                        step.setBearbeitungsstatusEnum(StepStatus.OPEN);
                    }
                }

                return true;
            }
        }).anyTimes();

        PowerMock.replay(CloseStepHelper.class);

        Prefs prefs = new Prefs();
        prefs.loadPrefs(resourcesFolder + "ruleset.xml");
        Fileformat ff = new MetsMods(prefs);
        ff.read(metaTarget.toString());

        PowerMock.mockStatic(MetadatenHelper.class);
        EasyMock.expect(MetadatenHelper.getMetaFileType(EasyMock.anyString())).andReturn("mets").anyTimes();
        EasyMock.expect(MetadatenHelper.getFileformatByName(EasyMock.anyString(), EasyMock.anyObject())).andReturn(ff).anyTimes();

        PowerMock.mockStatic(StepManager.class);
        PowerMock.mockStatic(ProcessManager.class);
        PowerMock.replay(ConfigPlugins.class);
        PowerMock.replay(ConfigurationHelper.class);
        PowerMock.replay(MetadatenHelper.class);
        process = getProcess();
        EasyMock.expect(ProcessManager.getProcessById(EasyMock.anyInt())).andReturn(process).anyTimes();
        PowerMock.replay(ProcessManager.class);

    }

    @Test
    public void testConstructor() {
        ProjectExportPlugin plugin = new ProjectExportPlugin();
        assertNotNull(plugin);
    }

    @Test
    public void testConfigurationFile() {
        ProjectExportPlugin plugin = new ProjectExportPlugin();
        plugin.setTestDatabase(true);
        plugin.setProjectName("SampleProject");

        assertEquals("closed step", plugin.getFinishStepName());
    }

    @Test
    public void testPrepareExport() {
        ProjectExportPlugin plugin = new ProjectExportPlugin();
        plugin.setTestDatabase(true);
        plugin.setExportFolder(metadataDirectory.getAbsolutePath());
        List<Process> testProcesses = new ArrayList<>();
        testProcesses.add(process);

        Process p2 = createAdditionalTestProcess(2, "RM0166F05-0000001");
        Process p3 = createAdditionalTestProcess(3, "RM0166F05-0000004");
        Process p4 = createAdditionalTestProcess(4, "RM0166F05-0000019");

        testProcesses.add(p2);
        testProcesses.add(p3);
        testProcesses.add(p4);

        plugin.setTestList(testProcesses);

        plugin.setProjectName("SampleProject");

        plugin.prepareExport();

        Path excelFile = Paths.get(metadataDirectory.getAbsolutePath(), "metadata.xlsx");
        assertTrue(Files.exists(excelFile));

        Path imageFolder = Paths.get(metadataDirectory.getAbsolutePath(), "990012587030205171");
        assertTrue(Files.exists(imageFolder));

        List<String> fileNames = new ArrayList<>();
        try (DirectoryStream<Path> directoryStream = Files.newDirectoryStream(imageFolder)) {
            for (Path path : directoryStream) {
                if (!path.getFileName().toString().startsWith(".")) {
                    fileNames.add(path.getFileName().toString());
                }
            }
        } catch (IOException ex) {
        }
        Collections.sort(fileNames);
        assertEquals("RM0166F05-0000001_001.jpg", fileNames.get(0));
        assertEquals("RM0166F05-0000001_016.jpg", fileNames.get(15));

        for (Step step : process.getSchritte()) {
            if (step.getTitel().equals("test step to close")) {
                assertEquals(StepStatus.DONE, step.getBearbeitungsstatusEnum());
            } else if (step.getTitel().equals("locked step that should open")) {
                assertEquals(StepStatus.OPEN, step.getBearbeitungsstatusEnum());
            }
        }
    }

    private Process createAdditionalTestProcess(int id, String title) {
        Process p = new Process();
        p.setId(id);
        p.setTitel(title);
        p.setProjekt(process.getProjekt());
        List<Step> steps = createSteps(p);

        p.setSchritte(steps);
        try {
            File processDirectory = new File(metadataDirectory + File.separator + id);
            processDirectory.mkdirs();
            Path metaSource = Paths.get(resourcesFolder + "meta" + id + ".xml");
            Path metaTarget = Paths.get(processDirectory.getAbsolutePath(), "meta.xml");
            Files.copy(metaSource, metaTarget);
            createProcessDirectory(processDirectory);
        } catch (IOException e) {
        }
        List<Processproperty> properties = createProperties();

        p.setEigenschaften(properties);
        return p;
    }

    public Process getProcess() {
        Project project = createProject();

        Process process = new Process();
        process.setTitel("RM0166F01-0000001");
        process.setProjekt(project);
        process.setId(1);
        List<Step> steps = createSteps(process);

        process.setSchritte(steps);
        try {
            createProcessDirectory(processDirectory);
        } catch (IOException e) {
        }
        List<Processproperty> properties = createProperties();

        properties.add(createProperty("notes_01", "Manṭovah :  Be-vet Yehudah Shemuʼel mi-Prushah u-veno,   [386] 1626."));

        process.setEigenschaften(properties);
        return process;
    }

    private List<Processproperty> createProperties() {
        List<Processproperty> properties = new ArrayList<>();

        properties.add(createProperty("Marginalia", "N"));
        properties.add(createProperty("Censorship", "N"));
        properties.add(createProperty("Provenance", "Y"));
        properties.add(createProperty("Book is important", "yes"));
        properties.add(createProperty("Reason for insignificance", ""));
        properties.add(createProperty("NLI identifier", "990012587030205171"));
        properties.add(createProperty("Reason for missing NLI identifier", ""));
        properties.add(createProperty("OCLC identifier", "319712882"));


        return properties;
    }

    private Project createProject() {
        Project project = new Project();
        project.setTitel("SampleProject");
        project.setMetsRightsOwner("Centro Bibliografico");
        project.setMetsRightsOwnerSite("http://ucei.it/centro-bibliografico/");
        return project;
    }

    private List<Step> createSteps(Process process) {
        List<Step> steps = new ArrayList<>();
        Step s1 = new Step();
        s1.setReihenfolge(1);
        s1.setProzess(process);
        s1.setTitel("closed step");
        s1.setBearbeitungsstatusEnum(StepStatus.DONE);
        User user = new User();
        user.setVorname("Firstname");
        user.setNachname("Lastname");
        s1.setBearbeitungsbenutzer(user);
        steps.add(s1);

        Step s2 = new Step();
        s2.setReihenfolge(2);
        s2.setProzess(process);
        s2.setTitel("test step to close");
        s2.setBearbeitungsstatusEnum(StepStatus.OPEN);
        steps.add(s2);

        Step s3 = new Step();
        s3.setReihenfolge(3);
        s3.setProzess(process);
        s3.setTitel("locked step that should open");
        s3.setBearbeitungsstatusEnum(StepStatus.LOCKED);
        steps.add(s3);
        return steps;
    }

    private Processproperty createProperty(String name, String value) {
        Processproperty prop = new Processproperty();
        prop.setProzess(process);
        prop.setTitel(name);
        prop.setWert(value);
        prop.setType(PropertyType.String);
        return prop;
    }

    private XMLConfiguration getConfig() throws Exception {
        XMLConfiguration config = new XMLConfiguration("plugin_intranda_workflow_projectexport.xml");
        config.setListDelimiter('&');
        config.setReloadingStrategy(new FileChangedReloadingStrategy());
        return config;

    }

    private void createProcessDirectory(File processDirectory) throws IOException {

        // image folder
        File imageDirectory = new File(processDirectory.getAbsolutePath(), "images");
        imageDirectory.mkdir();
        // master folder
        File masterDirectory = new File(imageDirectory.getAbsolutePath(), "master_RM0166F05-0000001_media");
        masterDirectory.mkdir();
        for (int i = 1; i <= 16; i++) {
            createFile(masterDirectory, i);
        }

        // media folder
        File mediaDirectory = new File(imageDirectory.getAbsolutePath(), "RM0166F05-0000001_media");
        mediaDirectory.mkdir();
        for (int i = 1; i <= 16; i++) {
            createFile(mediaDirectory, i);
        }

    }

    private void createFile(File folder, int counter) throws IOException {
        String imagename;
        if (counter > 9) {
            imagename = "RM0166F05-0000001_0" + counter + ".jpg";
        } else {
            imagename = "RM0166F05-0000001_00" + counter + ".jpg";
        }
        File image = new File(folder.getAbsolutePath(), imagename);
        image.createNewFile();
    }
}
