package de.intranda.goobi.plugins;

import java.util.List;

import org.apache.commons.configuration.HierarchicalConfiguration;
import org.apache.commons.configuration.XMLConfiguration;
import org.apache.commons.configuration.tree.xpath.XPathExpressionEngine;
import org.apache.commons.lang.StringUtils;
import org.goobi.production.enums.PluginType;
import org.goobi.production.plugin.interfaces.IWorkflowPlugin;

import de.sub.goobi.config.ConfigPlugins;
import de.sub.goobi.persistence.managers.ProjectManager;
import lombok.Getter;
import lombok.Setter;
import lombok.extern.log4j.Log4j2;
import net.xeoh.plugins.base.annotations.PluginImplementation;

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

}
