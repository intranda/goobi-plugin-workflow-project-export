package de.intranda.goobi.plugins;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.configuration.XMLConfiguration;
import org.apache.commons.configuration.reloading.FileChangedReloadingStrategy;
import org.easymock.EasyMock;
import org.goobi.beans.Process;
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
import de.sub.goobi.helper.enums.StepStatus;
import de.sub.goobi.persistence.managers.ProcessManager;
import de.sub.goobi.persistence.managers.PropertyManager;
import de.sub.goobi.persistence.managers.StepManager;

@RunWith(PowerMockRunner.class)
@PrepareForTest({ ConfigPlugins.class, StepManager.class, ConfigurationHelper.class, ProcessManager.class, PropertyManager.class })
@PowerMockIgnore({ "javax.management.*" })
public class ProjectExportPluginTest {

    @Rule
    public TemporaryFolder folder = new TemporaryFolder();
    private File processDirectory;
    private File metadataDirectory;
    private Process process;

    @Before
    public void setUp() throws Exception {
        metadataDirectory = folder.newFolder("metadata");
        processDirectory = new File(metadataDirectory + File.separator + "1");
        processDirectory.mkdirs();
        String metadataDirectoryName = metadataDirectory.getAbsolutePath() + File.separator;

        XMLConfiguration config = getConfig();
        PowerMock.mockStatic(ConfigPlugins.class);
        EasyMock.expect(ConfigPlugins.getPluginConfig("intranda_workflow_projectexport")).andReturn(config).anyTimes();
        PowerMock.replay(ConfigPlugins.class);

        PowerMock.mockStatic(ConfigurationHelper.class);
        ConfigurationHelper configurationHelper = EasyMock.createMock(ConfigurationHelper.class);
        EasyMock.expect(ConfigurationHelper.getInstance()).andReturn(configurationHelper).anyTimes();
        EasyMock.expect(configurationHelper.getMetsEditorLockingTime()).andReturn(1800000l).anyTimes();
        EasyMock.expect(configurationHelper.isAllowWhitespacesInFolder()).andReturn(false).anyTimes();
        EasyMock.expect(configurationHelper.useS3()).andReturn(false).anyTimes();
        EasyMock.expect(configurationHelper.isUseMasterDirectory()).andReturn(true).anyTimes();
        EasyMock.expect(configurationHelper.isCreateMasterDirectory()).andReturn(false).anyTimes();
        EasyMock.expect(configurationHelper.isCreateSourceFolder()).andReturn(false).anyTimes();
        EasyMock.expect(configurationHelper.getMediaDirectorySuffix()).andReturn("media").anyTimes();
        EasyMock.expect(configurationHelper.getMasterDirectoryPrefix()).andReturn("master").anyTimes();
        EasyMock.expect(configurationHelper.getFolderForInternalProcesslogFiles()).andReturn("intern").anyTimes();
        EasyMock.expect(configurationHelper.getMetadataFolder()).andReturn(metadataDirectoryName).anyTimes();

        EasyMock.expect(configurationHelper.getScriptCreateDirMeta()).andReturn("").anyTimes();
        EasyMock.replay(configurationHelper);

        PowerMock.mockStatic(StepManager.class);
        PowerMock.mockStatic(ProcessManager.class);
        PowerMock.mockStatic(PropertyManager.class);
        PowerMock.replay(ConfigPlugins.class);
        PowerMock.replay(ConfigurationHelper.class);

        process = getProcess();

    }

    private void createProcessDirectory(boolean createSourceDirectory, boolean createThumbsDirectory, boolean createOcrDirectory) throws IOException {

        // image folder
        File imageDirectory = new File(processDirectory.getAbsolutePath(), "images");
        imageDirectory.mkdir();
        // master folder
        File masterDirectory = new File(imageDirectory.getAbsolutePath(), "master_fixture_media");
        masterDirectory.mkdir();
        File masterImageFile = new File(masterDirectory.getAbsolutePath(), "0001.tif");
        masterImageFile.createNewFile();
        // media folder
        File mediaDirectory = new File(imageDirectory.getAbsolutePath(), "fixture_media");
        mediaDirectory.mkdir();
        File mediaImageFile = new File(mediaDirectory.getAbsolutePath(), "0001.tif");
        mediaImageFile.createNewFile();
        if (createSourceDirectory) {
            // source folder
            File sourceDirectory = new File(imageDirectory.getAbsolutePath(), "fixture_source");
            sourceDirectory.mkdir();
            File sourceFile = new File(sourceDirectory.getAbsolutePath(), "fixture.zip");
            sourceFile.createNewFile();
        }
        // thumbs
        if (createThumbsDirectory) {
            File thumbsDirectory = new File(processDirectory.getAbsolutePath(), "thumbs");
            thumbsDirectory.mkdir();
            File thumbsMediaDirectory = new File(thumbsDirectory.getAbsolutePath(), "fixture_media_800");
            thumbsMediaDirectory.mkdir();
            File thumbsImageFile = new File(thumbsMediaDirectory.getAbsolutePath(), "0001.tif");
            thumbsImageFile.createNewFile();
        }
        // ocr
        if (createOcrDirectory) {
            File ocr = new File(processDirectory.getAbsolutePath(), "ocr");
            ocr.mkdir();
            File altoDirectory = new File(ocr.getAbsolutePath(), "fixture_alto");
            altoDirectory.mkdir();
            File altoFile = new File(altoDirectory.getAbsolutePath(), "0001.xml");
            altoFile.createNewFile();
        }
    }

    @Test
    public void testConstructor() {
        ProjectExportPlugin plugin = new ProjectExportPlugin();
        assertNotNull(plugin);
    }

    @Test
    public void testConfigurationFile() {
        ProjectExportPlugin plugin = new ProjectExportPlugin();
        plugin.setProjectName("SampleProject");

        assertEquals("Export with PREMIS data", plugin.getFinishStepName());

    }


    public Process getProcess() {
        Project project = new Project();
        project.setTitel("SampleProject");

        Process process = new Process();
        process.setTitel("fixture");
        process.setProjekt(project);
        process.setId(1);
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
        s2.setTitel("Image deletion step");
        s2.setStepPlugin("intranda_step_imagedeletion");
        s2.setBearbeitungsstatusEnum(StepStatus.OPEN);
        s2.setTypAutomatisch(true);
        steps.add(s2);

        Step s3 = new Step();
        s3.setReihenfolge(3);
        s3.setProzess(process);
        s3.setTitel("test step to deactivate");
        s3.setBearbeitungsstatusEnum(StepStatus.LOCKED);
        steps.add(s3);

        process.setSchritte(steps);

        return process;
    }

    private XMLConfiguration getConfig() throws Exception {
        XMLConfiguration config = new XMLConfiguration("plugin_intranda_workflow_projectexport.xml");
        config.setListDelimiter('&');
        config.setReloadingStrategy(new FileChangedReloadingStrategy());
        return config;

    }
}
