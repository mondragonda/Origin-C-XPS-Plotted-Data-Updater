#include <Origin.h>
#include <GetNbox.h>
#include "importer.h"


test()
{
    XPSGraphExperimentUpdater xpsGraphUpdater();
}

class XPSGraphExperimentUpdater 
{
    public:
        XPSGraphExperimentUpdater();
    private:
    	XPSExperimentImporter xpsExperimentImporter;
        void displayFilePickerDialog();
        void displayConfirmationDialog(Tree);
        void displayWarningDialog(Tree);
        void displaySucceedDialog();
        void importUpdatedData(string);
        void rebuildGraphs();
        void destroyTempBooks();
        bool buttonevent(TreeNode& ,int,int,Dialog&);
};

XPSGraphExperimentUpdater::XPSGraphExperimentUpdater()
{
	displayFilePickerDialog();
}


void XPSGraphExperimentUpdater::displayFilePickerDialog()
{
    GETN_TREE(testTree)
    GETN_BUTTON(Path, "Seleccione el archivo con la actualización", xpsExperimentImporter.getExperimentFileName());
    GETN_OPTION_EVENT(buttonEvent) 
    
    if(GetNBox(testTree, NULL, NULL, NULL, NULL) == 1)
    	displayConfirmationDialog(testTree);
        
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

void XPSGraphExperimentUpdater::displayConfirmationDialog(Tree testTree)
{
	string confirmationDialogTitle = "Actualizar";
	string confirmationDialogMessage = "¿Esta seguro de actualizar el experimento?";
	
	int buttonTypeValue = MB_YESNO|MB_ICONQUESTION;
	
	int buttonClicked = MessageBox(GetWindow(), confirmationDialogMessage, confirmationDialogTitle, buttonTypeValue);
	
	if(buttonClicked == IDYES)
	{
		displayWarningDialog(testTree);;
	}
	
}

void XPSGraphExperimentUpdater::displayWarningDialog(Tree testTree)
{
	string warningDialogTitle = "Actualizar";
	string warningDialogMessage = "¡Los cambios no podran deshacerse!";
	
	int buttonTypeValue = MB_OKCANCEL|MB_ICONEXCLAMATION;
	
	if(MessageBox(GetWindow(),warningDialogMessage, warningDialogTitle, buttonTypeValue) == IDOK)
	  {
	  	    importUpdatedData(testTree.Path.strVal);
	 		rebuildGraphs();
	 		displaySucceedDialog();
	  }
}


void XPSGraphExperimentUpdater::rebuildGraphs()
{
	int amountOfBooks = Project.WorksheetPages.Count() / 2;
	
	vector<string> booksNames;
	booksNames = xpsExperimentImporter.getExperimentsNames();
	
	Dataset dataSet;
	
	for(int bookIndex=0; bookIndex < amountOfBooks; bookIndex++)
	{
  		string bookName = booksNames[bookIndex];
		WorksheetPage worksheetPage(bookName);
		WorksheetPage updatedWorksheetPage("update"+bookName);
		
		//try{
			for(int layerIndex=0; layerIndex < worksheetPage.Layers.Count(); layerIndex++)
				{
					Worksheet worksheet = worksheetPage.Layers(layerIndex);
					Worksheet updatedWorksheet = updatedWorksheetPage.Layers(layerIndex);
			
					for(int columnIndex=0; columnIndex < updatedWorksheet.GetNumCols(); columnIndex++)
						{
		    	
							Dataset dataSource(updatedWorksheet, columnIndex);
							
							
							if(worksheet.Columns(columnIndex) == NULL)
							{
								int newColumnIndex = worksheet.AddCol();
								
								Dataset dataSet(worksheet, newColumnIndex);
								
								worksheet.Columns(newColumnIndex).SetLongName(updatedWorksheet.Columns(columnIndex).GetLongName());
								worksheet.Columns(newColumnIndex).SetUnits(updatedWorksheet.Columns(columnIndex).GetUnits());
							}
							
							if(worksheet.Columns(columnIndex).GetLongName() != updatedWorksheet.Columns(columnIndex).GetLongName())
							{
								string newColumnName;
								worksheet.InsertCol(columnIndex, updatedWorksheet.Columns(columnIndex).GetLongName(), newColumnName, false);
								
							}
							
							Dataset dataSet(worksheet, columnIndex);
							
							
							dataSet = dataSource;
							
							
							dataSet.Update(FALSE, REDRAW_REALTIME_SCOPE);
		    	
						}
				}
		//}catch(int exceptionCode){
			//destroyTempBooks();
		//}

	}			
	
	//destroyTempBooks();
}

void XPSGraphExperimentUpdater::destroyTempBooks()
{
	foreach(WorksheetPage worksheetPage in Project.WorksheetPages)
	{
		if(worksheetPage.GetName().Find("update") > -1)
		     worksheetPage.Destroy();
	}
}

void XPSGraphExperimentUpdater::displaySucceedDialog()
{
	MessageBox(GetWindow(), "Los datos se han actualizado.", "Exito", MB_OK);
}

void XPSGraphExperimentUpdater::importUpdatedData(string updatedFilePath)
{
	xpsExperimentImporter.importExperiment(updatedFilePath);
}

