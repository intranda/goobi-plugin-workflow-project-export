package de.intranda.goobi.plugins;

import java.util.List;

import org.apache.commons.lang.StringUtils;
import org.goobi.production.enums.PluginType;
import org.goobi.production.plugin.interfaces.IWorkflowPlugin;

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
    private String selectedProjectName;

    @Setter
    private List<String> allProjectNames = null;






    public List<String> getAllProjectNames() {
        if (allProjectNames == null) {
            allProjectNames = ProjectManager.getAllProjectTitles(true);
        }
        return allProjectNames;
    }



    public void setSelectedProjectName(String selectedProjectName) {
        if (StringUtils.isBlank(this.selectedProjectName) || !this.selectedProjectName.equals(selectedProjectName)) {
            this.selectedProjectName = selectedProjectName;
            getConfig();
        }
    }

    private void getConfig() {

    }

}
