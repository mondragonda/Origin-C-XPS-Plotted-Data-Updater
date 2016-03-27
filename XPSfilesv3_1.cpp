/*------------------------------------------------------------------------------*
 * File Name:	XPSfiles														*
 * Creation: 	9.8.15															*
 * Purpose: 	Provide a class to handle CASA export files						*
 * 				Cabrera Ortega Jesus Efren										*
 * 				efren.cabrera@uabc.edu.mx										*
 *------------------------------------------------------------------------------*/
 

#include <Origin.h>
#include "importer.h"


main()
{
	string experimentFilePath = GetOpenBox("*.txt");
	
	XPSExperimentImporter xpsExperimentImporter();
	
	xpsExperimentImporter.importExperiment(experimentFilePath);
	
	vector<string> experimentNames;
	experimentNames = xpsExperimentImporter.getExperimentsNames();
	
	for(int i = 0; i < xpsExperimentImporter.getNumberOfExperiments(); i++)
		out_str(experimentNames[i]);
	
	out_int("Number of experiments: ", xpsExperimentImporter.getNumberOfExperiments());    
	
   
}

