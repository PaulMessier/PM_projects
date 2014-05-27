//"Stack>ROI>Measure"
//This macro is designed to create RIO based on user specifications for stacked images, ROI are then measured producing a tabulation of results
//Original code by Dr. Lai Ding, Harvard NeuroDiscovery Center Enhanced Neuroimaging Core 
//Designed by Maggie Wessling and Paul Messier, Boston 2014
//version 3

//add roiarrange function get rid of Anomaly roi.
 
//add user choose 3 slices or 2 slice stack
//add user define circularity parameter and threshold adjust.

// set front color as white and black color as black. 
// this prevent problem while doing thresholding.
run("Colors...", "foreground=white background=black selection=yellow");
run("Options...", "iterations=1 count=1 black edm=Overwrite");

// global variables.
var threshold_method, spot_size_min, spot_size_max, pixelscale, pixelunit, threshold_factor, mincir, slicecount;
var spot_size=newArray(10000), spot_class=newArray(10000), spot_mean_t1=newArray(10000), spot_mean_t2=newArray(10000), spot_mean_t3=newArray(10000), spot_count;
var Spotcount=newArray(500), Increase=newArray(500), Decrease=newArray(500), Nochange=newArray(500), Anomaly=newArray(500);
var percent, spotclass;
var f, flag=0;

// user input raw and result folder.  raw folder should only contain aligned stack images (each stack with 3 images.
rawdir=getDirectory("Choose Raw Data Folder");
resultdir=getDirectory("Choose Result Data Folder");

//input parameters
parameterinput();

// get file list in raw data folder.
list=getFileList(rawdir);

//for every file in raw folder.
for(f=0;f<list.length;f++)
 {
  open(rawdir+list[f]);
  flag=0;

  //find spots and save roi file
  spotfind();

  //if spots find, then do the measuremetn and assign spot class
  if(flag==0)
   {
    measurement();
    classification();
    roiarrange();
   }
  run("Close All");
 }

//generate "summary.xls" file
summary();

function summary()
 {
  print("RawDataFolder	"+rawdir);
  print("ResultDataFolder	"+resultdir);
  print("Percent%	SLiceCount	ThresholdMethod	ThresholdFactor	SpotMinSize	SpotMaxSize	MinCircularity");
  print(percent	+"	"+slicecount+"	"+threshold_method+"	"+threshold_factor+"	"+spot_size_min+"	"+spot_size_max+"	"+mincir);
  print("Filename	TotalSpot#	Increase#	Decrease#	Nochange#	Anomaly#");
  for(f=0;f<list.length;f++)
   print(list[f]+"	"+Spotcount[f]+"	"+Increase[f]+"	"+Decrease[f]+"	"+Nochange[f]+"	"+Anomaly[f]);

  selectWindow("Log");
  saveAs("Text", resultdir+"Summary.xls");
  run("Close");
 	
 }

function spotfind()
 {
  run("8-bit");
  rename("raw");
  //setSlice(3);
  //run("Duplicate...", "title=test");
  //rename("test");

  //use max project - min project to generate mask file to identify spots
  run("Z Project...", "start=1 stop="+nSlices+" projection=[Max Intensity]"); rename("max");
  selectImage("raw");
  run("Z Project...", "start=1 stop="+nSlices+" projection=[Min Intensity]"); rename("min");
  imageCalculator("Subtract create", "max","min");
  rename("test");
  run("Median...", "radius=2");

  // apply threshold method user input in dialog window
  setAutoThreshold(threshold_method);
  //adjust threshold value by multiple threshold_factor
  getThreshold(lower,upper);
  setThreshold(lower*threshold_factor,upper);
  setOption("BlackBackground", true);
  run("Convert to Mask");

  // apply size limitation user input in dialoag window
  run("Analyze Particles...", "size="+spot_size_min+"-"+spot_size_max+" pixel circularity="+mincir+"-1.00 show=Nothing add");
  if( roiManager("Count") > 0 ) 
     roiManager("Save", resultdir+list[f]+"_ROI.zip");
  else flag=1;   
  close();
 }

function measurement()
 {
  selectWindow("raw");	
  getVoxelSize(pixelscale, tmp, tmp, pixelunit);
  run("Properties...", "channels=1 slices="+slicecount+" frames=1 unit=pixel pixel_width=1 pixel_height=1 voxel_depth=1 frame=[0 sec] origin=0,0");
  run("Set Measurements...", "area mean standard modal min center integrated stack display redirect=None decimal=8");
  run("Clear Results");

  //measure mean intensity for all slices
  setSlice(1); roiManager("Measure");
  setSlice(2); roiManager("Measure");
  if( slicecount == 3 )
    setSlice(3); roiManager("Measure");

  selectWindow("Results");
  saveAs("Text", resultdir+list[f]+"_rawresults.xls");
  spot_count=roiManager("Count");
  roiManager("reset");

  for(i=0;i<spot_count;i++)
   {
    spot_size[i]= getResult("Area", i);
    spot_mean_t1[i]= getResult("Mean", i);
    spot_mean_t2[i]= getResult("Mean", i+spot_count);
    if( slicecount == 3 )
      spot_mean_t3[i]= getResult("Mean", i+spot_count*2);
   }
  run("Clear Results");
 }

function classification()
 {
  print("Filename	"+rawdir+list[f]);
  print("TotalSpot#	"+spot_count);
  print("Percent%	"+percent);
  p_up = 1 + percent/100;
  p_down = 1 - percent/100 ;
  if( slicecount == 3 )
    print("Spot#	MeanT1Norm	MeanT2Norm	MeanT3Norm	Class");
  if( slicecount == 2 )  
    print("Spot#	MeanT1Norm	MeanT2Norm	Class");
    
  for(i=0;i<spot_count;i++)
   {
    t1 = spot_mean_t1[i]; 
    t2 = spot_mean_t2[i]; 
    if( slicecount == 3 )
      t3 = spot_mean_t3[i]; 

    if( slicecount==2 )  classmethod2(t1,t2,p_up,p_down);
    if( slicecount==3 )  classmethod3(t1,t2,t3,p_up,p_down);
    
    if( spotclass == "A" ) { Anomaly[f]++; spot_class[i] = "A" ; }
    if( spotclass == "I" ) { Increase[f]++;  spot_class[i] = "I" ; }
    if( spotclass == "D" ) { Decrease[f]++;  spot_class[i] = "D" ; }
    if( spotclass == "N" )  { Nochange[f]++; spot_class[i] = "N" ; }
    Spotcount[f]=spot_count;
  
    if( slicecount == 3 )
       print(i+1+"	"+100+"	"+100*t2/t1+"	"+100*t3/t1+"	"+spotclass);
    if( slicecount == 2 )
       print(i+1+"	"+100+"	"+100*t2/t1+"	"+spotclass);        
   }
   
  selectWindow("Log");
  saveAs("Text", resultdir+list[f]+"_result.xls");
  run("Close");
 }

function roiarrange()
 {
  roiManager("Open", resultdir+list[f]+"_ROI.zip");
  k=0;
  for(i=0;i<spot_count;i++)
   {
    if( spot_class[i] == "A" )
     {
      roiManager("Select",i-k);
      roiManager("Delete");
      k++;
     }
   }
   roiManager("Save", resultdir+list[f]+"_ROInew.zip"); 
   roiManager("reset");
 }

function classmethod2(t1,t2,p_up,p_down)
 {
  if( t2 > t1*p_up )    // spots with increased intensity
      tmpclass="I";
  if( t2 < t1*p_down )   // spots with decreased intensity
      tmpclass="D";  
  if( (t2 >= t1*p_down) && (t2 <= t1*p_up) )   // spots with unchanged intensity
      tmpclass="N";
      
  spotclass = tmpclass;
 }

function classmethod3(t1,t2,t3,p_up,p_down)
 {	
  if( t3 > t1*p_up )    // spots with increased intensity
     {
      tmpclass="I";
      if( t2 < t1*p_down )  tmpclass="A";
      if( t2 > t3*p_up ) tmpclass="A";
     }
     
    if( t3 < t1*p_down )   // spots with decreased intensity
     {
      tmpclass="D";
      if( t2 > t1*p_up )  tmpclass="A";
      if( t2 < t3*p_down ) tmpclass="A";
     }
     
    if( (t3 >= t1*p_down) && (t3 <= t1*p_up) )   // spots with unchanged intensity
     {
      tmpclass="N";
      if( (t2 > maxOf(t1,t3)*p_up) || (t2 < minOf(t1,t3)*p_down) )  tmpclass="A";
     }
     
  spotclass = tmpclass;
 }

// read user input parameters, assign value to global variables 
function parameterinput()
 {
  Dialog.create("Parameters");
  Dialog.addChoice("Slice Count:", newArray(2, 3));
  Dialog.addChoice("Threshold Method:", newArray("Default dark", "Huang dark", "Intermodes dark", "IsoData dark", "IJ_IsoData dark", "Li dark", "MaxEntropy dark", "Mean dark", "MinError dark", "Minimum dark", "Moments dark", "Otsu dark", "Percentile dark", "RenyiEntropy dark", "Shanbhag dark", "Triangle dark", "Yen dark" ));
  Dialog.addNumber("Threhsold Factor:", 1);
  Dialog.addNumber("Minimum Spot Size(pixel):", 50);
  Dialog.addNumber("Maximum Spot Size(pixel):", 600);
  Dialog.addNumber("Min Circularity:", 0.3);
  Dialog.addNumber("Percent(%):", 30);
  
  Dialog.show();
  
  slicecount = Dialog.getChoice();
  threshold_method = Dialog.getChoice();
  threshold_factor = Dialog.getNumber();
  spot_size_min = Dialog.getNumber();
  spot_size_max = Dialog.getNumber();
  mincir = Dialog.getNumber();
  percent = Dialog.getNumber();

  slicecount=round(slicecount);
 }

 
