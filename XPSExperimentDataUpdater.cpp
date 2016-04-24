#include <Origin.h>
#include <GetNbox.h>
#include "importer.h"


test()
{
    XPSExperimentDataUpdater xpsExperimentDataUpdater();
}

class XPSExperimentDataUpdater 
{
    public:
        XPSExperimentDataUpdater();
    private:
    	Collection<WorksheetPage> projectWorksheetPages;
    	XPSExperimentImporter xpsExperimentImporter;
    	void loadProjectWorksheetPages();
        void displayUpdateFilePickerDialog();
        void displayUpdateConfirmationDialog(Tree);
        void displayUpdateWarningDialog(Tree);
        void displayUpdateSucceededDialog();
        void importUpdatedData(string);
        void updateExperimentData();
        void addData();
        bool worksheetPageNameDoesNotContainUpdateWord(string);
        WorksheetPage getWorksheetPageByName(string);
        bool updatedLayerExistsInOriginalWorksheetPage(Collection<Layer>, Layer);
        Layer getLayerByName(WorksheetPage, string);
        void copyUpdatedData(Worksheet,Worksheet);
        Layer getWorksheetPageLayerByLayerIndex(WorksheetPage, int);
        bool updatedColumnExistsInOriginalLayer(Column, Worksheet);
        Column getWorksheetColumnByColumnIndex(Worksheet, int);
        Column getOriginalColumn(Worksheet, string);
        void deleteData();
        bool originalLayerExistsInUpdatedWorksheetPage(WorksheetPage, Layer);
        bool originalColumnDoesNotExistInUpdatedLayer(Worksheet, Column);
        void destroyTempBooks();
        bool buttonevent(TreeNode& ,int,int,Dialog&);
};

XPSExperimentDataUpdater::XPSExperimentDataUpdater()
{
	loadProjectWorksheetPages();
	displayUpdateFilePickerDialog();
}

void XPSExperimentDataUpdater::loadProjectWorksheetPages()
{
	projectWorksheetPages = Project.WorksheetPages;
}

void XPSExperimentDataUpdater::displayUpdateFilePickerDialog()
{
    GETN_TREE(testTree)
    GETN_BUTTON(Path, "Seleccione el archivo con la actualización", xpsExperimentImporter.getExperimentFileName());
    GETN_OPTION_EVENT(buttonEvent) 
    
    if(GetNBox(testTree, NULL, NULL, NULL, NULL) == 1)
    	displayUpdateConfirmationDialog(testTree);
        
}


bool buttonEvent(TreeNode& myTree, int nRow, int nType, Dialog& theDlg)
{
	if(TRGP_STR_BUTTON == nType && nRow >= 0)
	{
		string strPath = GetOpenBox("*.txt");
		
		myTree.Path.strVal = strPath;
		
		return true;
	}
	else
    {
    	return false;
	}
		
}

void XPSExperimentDataUpdater::displayUpdateConfirmationDialog(Tree testTree)
{
	string confirmationDialogTitle = "Actualizar";
	string confirmationDialogMessage = "¿Esta seguro de actualizar el experimento?";
	
	int buttonTypeValue = MB_YESNO|MB_ICONQUESTION;
	
	int buttonClicked = MessageBox(GetWindow(), confirmationDialogMessage, confirmationDialogTitle, buttonTypeValue);
	
	if(buttonClicked == IDYES)
	{
		displayUpdateWarningDialog(testTree);;
	}
	
}

void XPSExperimentDataUpdater::displayUpdateWarningDialog(Tree testTree)
{
	string warningDialogTitle = "Actualizar";
	string warningDialogMessage = "¡Los cambios no podran deshacerse!";
	
	int buttonTypeValue = MB_OKCANCEL|MB_ICONEXCLAMATION;
	
	if(MessageBox(GetWindow(),warningDialogMessage, warningDialogTitle, buttonTypeValue) == IDOK)
	  {
	  	    importUpdatedData(testTree.Path.strVal);
	 		updateExperimentData();
	 		displayUpdateSucceededDialog();
	  }
}

void XPSExperimentDataUpdater::importUpdatedData(string updatedFilePath)
{
	xpsExperimentImporter.importExperiment(updatedFilePath);
}


void XPSExperimentDataUpdater::updateExperimentData()
{
	addData();
	deleteData();
	destroyTempBooks();

}

void XPSExperimentDataUpdater::addData()
{
	
	foreach(WorksheetPage worksheetPage in projectWorksheetPages)
	{
		string worksheetPageName = worksheetPage.GetLongName();
		
		if(worksheetPageNameDoesNotContainUpdateWord(worksheetPageName))
		{
			WorksheetPage updatedWorksheetPage = getWorksheetPageByName("update"+worksheetPageName);
			
			if(updatedWorksheetPage != NULL)
			{
				Collection<Layer> originalWorksheetPageLayers;
				originalWorksheetPageLayers = worksheetPage.Layers;
			
				Collection<Layer> updatedWorksheetPageLayers;
				updatedWorksheetPageLayers= updatedWorksheetPage.Layers;
			
				foreach(Layer updatedLayer in updatedWorksheetPageLayers)
				{
					if(updatedLayerExistsInOriginalWorksheetPage(originalWorksheetPageLayers, updatedLayer))
					{
						string layerName = updatedLayer.GetName();
					
						Layer originalLayer = getLayerByName(worksheetPage, layerName);
					
						copyUpdatedData(updatedLayer, originalLayer);
					}
					else
					{
						int newLayerIndex = worksheetPage.AddLayer();
					
						Worksheet newWorksheet = getWorksheetPageLayerByLayerIndex(worksheetPage, newLayerIndex);
						
						newWorksheet.SetName(updatedLayer.GetName());
					
						copyUpdatedData(updatedLayer, newWorksheet);
					}
			
				}
				
			}
			
		}
	}
	
}

bool XPSExperimentDataUpdater::worksheetPageNameDoesNotContainUpdateWord(string worksheetPageLongName)
{
	if(worksheetPageLongName.Find("update") <= -1)
		return true;
	else
		return false;
}

WorksheetPage XPSExperimentDataUpdater::getWorksheetPageByName(string worksheetPageName)
{
	foreach(WorksheetPage worksheetPage in projectWorksheetPages)
	{
		if(  worksheetPage.GetLongName() == worksheetPageName
		   ||worksheetPage.GetName() == worksheetPageName)
			{
				return worksheetPage;
				
				break;
			}
	}
	
	return NULL;
}

bool XPSExperimentDataUpdater::updatedLayerExistsInOriginalWorksheetPage(Collection<Layer> originalWorksheetPageLayers, Layer updatedLayer)
{
	foreach(Layer originalLayer in originalWorksheetPageLayers)
	{
		if(originalLayer.GetName() == updatedLayer.GetName())
		{
			return true;
			
			break;
		}
	}
	
	return false;
	
}

Layer XPSExperimentDataUpdater::getLayerByName(WorksheetPage worksheetPage, string layerName)
{
	foreach(Layer layer in worksheetPage.Layers)
	{
		if(layer.GetName() == layerName)
		{
			return layer;
			
			break;
		}
	}
	
	return NULL;
}

Layer XPSExperimentDataUpdater::getWorksheetPageLayerByLayerIndex(WorksheetPage worksheetPage, int layerIndex)
{
	return worksheetPage.Layers(layerIndex);
}

void XPSExperimentDataUpdater::copyUpdatedData(Worksheet updatedWorksheet, Worksheet originalWorksheet)
{
	int updatedColumnIndex = 0;
	
	foreach(Column updatedColumn in updatedWorksheet.Columns)
	{
		Dataset dataSource(updatedColumn);
		
		if(updatedColumnExistsInOriginalLayer(updatedColumn, originalWorksheet))
		{
			string columnName = updatedColumn.GetLongName();
			
			Column originalColumn = getOriginalColumn(originalWorksheet, columnName);
			
			Dataset dataSet(originalColumn);
			
			originalColumn = updatedColumn;
			
			dataSet.Update(FALSE, REDRAW_REALTIME_SCOPE);
		}
		else
		{	
			Column originalColumn = originalWorksheet.Columns(updatedColumnIndex);
		
			if(originalColumn == NULL)
			{
				int newColumnIndex = originalWorksheet.AddCol();
			
				Column newColumn = getWorksheetColumnByColumnIndex(originalWorksheet, newColumnIndex);
			
				Dataset dataSet(newColumn);
			
				newColumn.SetLongName(updatedColumn.GetLongName());
			
				newColumn.SetUnits(updatedColumn.GetUnits());
			
				dataSet = dataSource;
			}
			else
			{
				Dataset dataSet(originalColumn);
							
				originalColumn.SetLongName(updatedColumn.GetLongName());
			
				originalColumn.SetUnits(updatedColumn.GetUnits());
			
				dataSet = dataSource;
				
			}
				
		}
		
		updatedColumnIndex++;
	}
		
}

bool XPSExperimentDataUpdater::updatedColumnExistsInOriginalLayer(Column updatedColumn, Worksheet originalWorksheet)
{
	foreach(Column column in originalWorksheet.Columns)
	{
		if(column.GetLongName() == updatedColumn.GetLongName())
		{
			return true;
			
			break;
		}
	}
	
	return false;
}

Column XPSExperimentDataUpdater::getWorksheetColumnByColumnIndex(Worksheet worksheet, int columnIndex)
{
	return worksheet.Columns(columnIndex);
}

Column XPSExperimentDataUpdater::getOriginalColumn(Worksheet originalWorksheet, string columnName)
{
	foreach(Column column in originalWorksheet.Columns)
	{
		if(column.GetLongName() == columnName)
		{
			return column;
			
			break;
		}
	}
	
	return NULL;
	
}

void XPSExperimentDataUpdater::deleteData()
{
	foreach(WorksheetPage worksheetPage in projectWorksheetPages)
	{
		string worksheetPageName = worksheetPage.GetLongName();
		
		if(worksheetPageNameDoesNotContainUpdateWord(worksheetPageName))
		{
			WorksheetPage updatedWorksheetPage = getWorksheetPageByName("update"+worksheetPageName);
			
			if(updatedWorksheetPage != NULL)
			{
				Collection<Layer> originalWorksheetPageLayers;
				originalWorksheetPageLayers = worksheetPage.Layers;
			
				Collection<Layer> updatedWorksheetPageLayers;
				updatedWorksheetPageLayers= updatedWorksheetPage.Layers;
			
				foreach(Layer originalLayer in originalWorksheetPageLayers)
				{
					if(originalLayerExistsInUpdatedWorksheetPage(updatedWorksheetPage, originalLayer))
					{
						string layerName = originalLayer.GetName();
					
						Worksheet originalWorksheet = getLayerByName(worksheetPage, layerName);
						Worksheet updatedWorksheet = getLayerByName(updatedWorksheetPage, layerName);
					
						foreach(Column column in originalWorksheet.Columns)
						{
							if(originalColumnDoesNotExistInUpdatedLayer(updatedWorksheet, column))
								column.Destroy();
						}
					}
					else
					{
						originalLayer.Destroy();
					
					}
				}
			}
		}
	}
	
}

bool XPSExperimentDataUpdater::originalLayerExistsInUpdatedWorksheetPage(WorksheetPage updatedWorksheetPage, Layer originalLayer)
{
	foreach(Layer layer in updatedWorksheetPage.Layers)
	{
		if(layer.GetName() == originalLayer.GetName())
		{
			return true;
			
			break;
		}
	}
	
	return false;
}


bool XPSExperimentDataUpdater::originalColumnDoesNotExistInUpdatedLayer(Worksheet updatedWorksheet, Column originalColumn)
{
	foreach(Column column in updatedWorksheet.Columns)
	{
		if(column.GetLongName() == originalColumn.GetLongName())
		{
			return false;
			
			break;
		}
	}
	
	return true;
}


void XPSExperimentDataUpdater::destroyTempBooks()
{
	foreach(WorksheetPage worksheetPage in Project.WorksheetPages)
	{
		if(worksheetPage.GetName().Find("update") > -1)
		     worksheetPage.Destroy();
	}
}

void XPSExperimentDataUpdater::displayUpdateSucceededDialog()
{
	MessageBox(GetWindow(), "Los datos se han actualizado.", "Exito", MB_OK);
}

