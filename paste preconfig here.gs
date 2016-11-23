function formMule_preconfig() {
// if you are interested in sharing your complete workflow system for others to copy (with script settings)
// Select the "Generate preconfig()" option in the menu and
//#######Paste preconfiguration code below before sharing your system for copy#######

  
  
  
  
    
//#######End preconfiguration code#######
// Note: By design, calendarID's will not be copied along with script settings. 
 //Fetch system name, if this script is part of a New Visions system
  var systemName = NVSL.getSystemName();
  if (systemName) {
    ScriptProperties.setProperty('systemName', systemName)
  }
  //Fetch institutional tracking code.  If it exists, launch initialize function (autolaunch step 1 for repeat users)
  //If it doesn't exist, the checkInstitutionalTrackingCode() will launch the tracking settings UI.
  var institutionalTrackingString = NVSL.checkInstitutionalTrackingCode();
}
