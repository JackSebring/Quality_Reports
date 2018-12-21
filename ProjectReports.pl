#given project sheet id find, project name, project ID, project location, & 4 pdf documents 

use warnings;
use strict;
use Cwd;
use File::Basename;
use Excel::Writer::XLSX;

my $ProjectSheetID;
my $vehicle_platform;
my @FolderLocations;
my $FolderLocation;
my $ReleaseNumber; #??????Correct name for each release version within forward relationship tab
my @array;
my $mks_command;
my $filename;
my $counter;
my $one_line;
my $project_id;
my $customer;
my @lines;
my $ForwardRel;
my @ForwardRels;
my @ReleaseSheetIDs;
my $ReleaseSheetID;
my @Unique_Components;
my @CompBug_RedGreenArray;
my @QAC_RedGreenArray;
my @CodeReview_RedGreenArray;
my @Components;
my $x;
my $y;
my @Folder_Array;
my @CodeReview_Array;
my @COSIPA_Array;
my @QAC_Array;
my @CompBug_Array;
my $SystemID;
my $SystemDesc;

print "Provide Project Sheet ID Please: ";
#$ProjectSheetID = <STDIN>;
$ProjectSheetID = "1356285"; #FORD
#$ProjectSheetID = "1630179"; #sr
#$ProjectSheetID = "1499487"; #FCA
#$ProjectSheetID = "1630180"; #sr2
#$ProjectSheetID = "1630179"; #v713 HEVHAD
#process Project Sheet ID and populate text file from MKS content

print "Getting information from Project Sheet ID ".$ProjectSheetID."\n";
$mks_command = `im setprefs --command=connect --nosave server.hostname=ffm-mks3`;
print $mks_command;
$mks_command = `im setprefs --command=connect --nosave server.port=7002`;
print $mks_command;
$mks_command = `im connect --hostname=ffm-mks3 --port=7002 --batch`;
print $mks_command;
$mks_command = `im viewissue $ProjectSheetID`;
print $mks_command;
open(LOG_FILE,'>01.project_information.txt');
print LOG_FILE $mks_command;
close(LOG_FILE);

#http://ffm-mks3:7001/si/viewrevision?projectName=d:/mks/archives/release/FIAT%5f520334%5fMKC1%5fESC%5fXJ2/17102400/04%2dApplicationSoftware/04%2dCodeAnalysis/BFU/BFU.pj&selection=17102400%5fBFU%5fstat%5fanalysis%5fwith%5fQAC.pdf

#take text file and search for lines that contain: Project Name, Project ID, File location, and forward relationships

print "PROCESSING STARTS HERE\n";
open(FILE, "<", "01.project_information.txt") or die("Can't open file");
@lines = <FILE>;
chomp(@lines);	
close(FILE);

foreach $one_line (@lines)    # Check if correct project sheet ID is used
{
	if($one_line =~ /(.*):\strue/){
		print $1."\n";
		push (@Components,$1);
	}
   if(index($one_line,"Type: Project Sheet") == 0)
    {
       print "Project Sheet found: ".$ProjectSheetID." \n";
    }
	
   if($one_line =~ /Release\sSheet\sApplication\s(.*):\s(.*)_(.*)_(.*)_(.*)_(.*)/){
		#Release Sheet application = $1???
		$customer = $2; #FORD
		$vehicle_platform = $3; #CD6
		$SystemID = $4; #MKC1
		$SystemDesc = $5;#ESC
		$ForwardRel = $6;#12345678
		print "Customer: ".$customer."\n";
		print "Vehicle Platform: ".$vehicle_platform."\n";
		print "SystemID: ".$SystemID."\n";
		print "System Desc: ".$SystemDesc."\n";
		print "Forward Relationship Number: ".$ForwardRel."\n";
		push @ForwardRels,$ForwardRel;
	}
   if($one_line =~ /ProjectID:\s(.*)/){
		$project_id = $1;
		print $project_id;
	}
}

foreach $x(@ForwardRels)
{
	push(@CodeReview_Array," ");
	push(@QAC_Array," ");
	push(@CompBug_Array," ");
	push(@Unique_Components," ");
	push(@CodeReview_RedGreenArray," ");
	push(@QAC_RedGreenArray," ");
	push(@CompBug_RedGreenArray," ");
	
	push (@FolderLocations, "d:/mks/archives/release/".$customer."_".$vehicle_platform."_".$SystemDesc."_".$project_id."/".$x."/".$x);
	$mks_command = `si setprefs --command=connect --nosave server.hostname=ffm-mks3`;
	$mks_command = `si setprefs --command=connect --nosave server.port=7001`;
	$mks_command = `si connect --hostname=ffm-mks3 --port=7001 --batch`;
	foreach $y (@Components)#http://ffm-mks3:7001/si/viewrevision?projectName=d:/mks/archives/release/FORD%5fCD6%5fMKC1%5fESC%5f2X2/18300314/04%2dApplicationSoftware/04%2dCodeAnalysis/WhDyAc/WhDyAc.pj&selection=18300314%5fWhDyAc%5freview%5frecord%5fcompilerbug.pdf
	{
		#print "http://ffm-mks3:7001/si/viewrevision?projectName=d:/mks/archives/release/".$customer."%5f".$vehicle_platform."%5f".$SystemID."%5f".$SystemDesc."%5f".$project_id."/".$x."/"."04%2dApplicationSoftware/04%2dCodeAnalysis/".$y."/". $y.".pj&selection=".$x."%5f".$y."%5fcode%5freview.pdf";
		push (@Unique_Components,$y);
		push(@CodeReview_Array,"http://ffm-mks3:7001/si/viewrevision?projectName=d:/mks/archives/release/".$customer."%5f".$vehicle_platform."%5f".$SystemID."%5f".$SystemDesc."%5f".$project_id."/".$x."/"."04%2dApplicationSoftware/04%2dCodeAnalysis/".$y."/". $y.".pj&selection=".$x."%5f".$y."%5fcode%5freview.pdf");
	
		push(@QAC_Array,"http://ffm-mks3:7001/si/viewrevision?projectName=d:/mks/archives/release/".$customer."%5f".$vehicle_platform."%5f".$SystemID."%5f".$SystemDesc."%5f".$project_id."/".$x."/"."04%2dApplicationSoftware/04%2dCodeAnalysis/".$y."/". $y.".pj&selection=".$x."%5f".$y."%5fstat%5fanalysis%5fwith%5fQAC.pdf");
		push(@CompBug_Array,"http://ffm-mks3:7001/si/viewrevision?projectName=d:/mks/archives/release/".$customer."%5f".$vehicle_platform."%5f".$SystemID."%5f".$SystemDesc."%5f".$project_id."/".$x."/"."04%2dApplicationSoftware/04%2dCodeAnalysis/".$y."/". $y.".pj&selection=".$x."%5f".$y."%5review%5frecord%5fcompilerbug.pdf");
		$mks_command = "si viewproject --project=d:/mks/archives/release/".$customer."_".$vehicle_platform."_".$SystemID."_".$SystemDesc."_".$project_id."/".$x."/04-ApplicationSoftware/04-CodeAnalysis/".$y."/".$y.".pj --filter=file:*_code_review.pdf"; # Call for entire folder
        $mks_command =`$mks_command`; 
		print "Exists: ".$mks_command."\n";
		if ($mks_command ne " " ){
			push(@CodeReview_RedGreenArray,$mks_command);
		}
		
		$mks_command = "si viewproject --project=d:/mks/archives/release/".$customer."_".$vehicle_platform."_".$SystemID."_".$SystemDesc."_".$project_id."/".$x."/04-ApplicationSoftware/04-CodeAnalysis/".$y."/".$y.".pj --filter=file:*_QAC.pdf"; # Call for entire folder
        $mks_command = `$mks_command`; 
		print "Exists: ".$mks_command."\n";
		if ($mks_command ne " " ){
			push(@QAC_RedGreenArray,$mks_command);
		}
		
		$mks_command = "si viewproject --project=d:/mks/archives/release/".$customer."_".$vehicle_platform."_".$SystemID."_".$SystemDesc."_".$project_id."/".$x."/04-ApplicationSoftware/04-CodeAnalysis/".$y."/".$y.".pj --filter=file:*_record_compilerbug.pdf"; # Call for entire folder
        $mks_command =`$mks_command`; 
		print "Exists: ".$mks_command."\n";
		if ($mks_command ne " " ){
			push(@CompBug_RedGreenArray,$mks_command);
		}
	}
	#												  d:/mks/archives/release/FORD%5fCD6%5fMKC1%5fESC%5f2X2/17300179/04%2dApplicationSoftware/02%2dReleaseVersion/controller/controller.pj
	#http://ffm-mks3:7001/si/viewrevision?projectName=d:/mks/archives/release/FORD%5fCD6%5fMKC1%5fESC%5f2X2/17300179/04%2dApplicationSoftware/02%2dReleaseVersion/controller/controller.pj&selection=2X2R008CKN0.cosipa.xlsx

	
					    #http://ffm-mks3:7001/si/viewrevision?projectName=d:/mks/archives/release/FORD%5fCD6%5fMKC1%5fESC%5f2X2/17300179/					04%2dApplicationSoftware/02%2dReleaseVersion/controller/controller.pj&selection=2X2R008CKN0.cosipa.xlsx
						#http://ffm-mks3:7001/si/viewrevision?projectName=d:/mks/archives/release/FORD_CD6_MKC1_ESC_2X2/17300819/							  04-ApplicationSoftware/02-ReleaseVersion/controller/controller.pj&selection=2X2R009W4MS.cosipa.xlsx
	$mks_command = "si viewproject --project=d:/mks/archives/release/".$customer."_".$vehicle_platform."_".$SystemID."_".$SystemDesc."_".$project_id."/".$x."/04-ApplicationSoftware/02-ReleaseVersion/controller/controller.pj --filter=file:*.cosipa.*";
	print "Please Type y\n";
	$mks_command = `$mks_command`;
	if ($mks_command eq "" or $mks_command eq " " or $mks_command eq "	" or $mks_command eq "\n"){
	$mks_command = "EMPTY";
	}
	if ($mks_command eq "DataOutput/DataOutput.pj subproject"){
	$mks_command = "EMPTY";
	}
	print "COSIPA: ".$mks_command."\n";
		# if ($mks_command ne " " )
		# {
		push(@COSIPA_Array,$mks_command);
			# if ($mks_command =~ /COSIPA:\s\sDataOutput\/DataOutput.pj\ssubproject\n\s(.*).cosipa.xlsx\sarchived\s(.*)/)
			# {
			# print "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\n";
				# $mks_command = $1;
				# push(@COSIPA_Array,$mks_command);
				
			# }
		#}
			
		#else {push(@COSIPA_Array, "EMPTY");}	
	#############################################################
	
	# my $query;	
	# my $create_query;
	# $mks_command = `si setprefs --command=connect --nosave server.hostname=ffm-mks3`;
	# print $mks_command;
	# $mks_command = `si setprefs --command=connect --nosave server.port=7001`;
	# print $mks_command;
	# $mks_command = `si connect --hostname=ffm-mks3 --port=7001 --batch`;
	# print $mks_command;
	
	
	#my $create_query = ''; 
	#$query = '[SWArch] (ProjectSheet ' . $ProjectSheetID . ') - All RSAs for ' . $vehicle_platform . ' (' . $project_id . ')';
	#$query = $ProjectSheetID.'RSAs for: '.$vehicle_platform.'	'.$project_id;
	#my $ProjectFieldQuery = "/Pool/System/MK C1/Component/Software";
	
	
	
	# $query = "im createquery --name=" .$ProjectSheetID.$vehicle_platform.$project_id. " --queryDefinition=\((field[".'"Summary"'."] contains ".$x." ))/";
	
	# print "################\n";
	# $create_query = `im viewquery "$query"`;
	# if (length($create_query) == 0) # If query doesn't exist, create one
	# {
	# print "@@@@@@@@@@@@@@@\n";
	# $create_query = `$query`;
	# }
	# else
	# {
	# print "%%%%%%%%%%%%%%%%%\n";
	# print "Query '" . $query . "' already exists. Creation of new Query not needed.\n" ;
	# }
	##################################################################
}
print "All COSIPA: \n";
my $b=0;
foreach (@COSIPA_Array){
print $b."	";
print $_."\n";
$b++;
}
#Take forward relationships and find version paths in summary ".......17300625"
#$mks_command = `im abouts`;

#Print data to Excel 

print "Excel File is now being written \n";
my $workbook  = Excel::Writer::XLSX->new( 'Reports.xlsx');
my $worksheet = $workbook->add_worksheet();
my $BreakFormat = $workbook->add_format();
$BreakFormat ->set_bg_color('black');
my $WhiteFormat = $workbook-> add_format();
$WhiteFormat ->set_bg_color('white');
my $StructureFormat= $workbook->add_format();

my $GreenFormat = $workbook->add_format();
$GreenFormat ->set_bg_color('#99ff66');
my $RedFormat = $workbook->add_format();
$RedFormat ->set_bg_color('#ff5050');
my $FrozenFormat = $workbook->add_format();
$FrozenFormat ->set_bg_color('#66ffff');

$worksheet->set_column('B:J',18, $StructureFormat);
$worksheet->set_row('1',7, $BreakFormat); 
$worksheet->set_row('3',7, $BreakFormat); 
$worksheet->set_column('A:A',7, $BreakFormat);
#$worksheet->set_column('M:AZ',7, $BreakFormat);
$worksheet->write("B1", "Continental AG");
$worksheet->write("C1", "SW Architecture");

$worksheet->write("B3", "Project Sheet ID");
$worksheet->write("C3",  "Customer");  	
$worksheet->write("D3", "Vehicle Platform");
$worksheet->write("E3", "Project ID");
$worksheet->write("F3", "Folder Location");
$worksheet->write("G3", "Release ID");
$worksheet->write("H3", "Component");
$worksheet->write("I3", "Release Version COSIPA");
$worksheet->write("J3", "Other COSIPA");
$worksheet->write("K3", "Code Review PDF");
$worksheet->write("L3",  "QAC Report");	
$worksheet->write("M3",  "Compiler Bug Report");

$worksheet->write("I1","Present",$GreenFormat);
$worksheet->write("J1","Not Present",$RedFormat);
$worksheet->write("K1","Frozen Members/Cancelled Release",$FrozenFormat);
my $rowcounter = 4;
my $colcounter = 10;
my $w;
my $z;

#code review array contains all possible code review pdf links
#code review redgreen array contains the archived 1.0+ files in MKS

foreach $z(@CodeReview_RedGreenArray){
		if ($z =~ /.*\.pdf\sarchived\s(.*)/)
		{		
			if ($1 == "1.0")
			{
				$worksheet->write($rowcounter,$colcounter, $CodeReview_Array[$rowcounter-4], $RedFormat);
				$rowcounter++;
			}
			else
			{
				$worksheet->write($rowcounter,$colcounter,$CodeReview_Array[$rowcounter-4], $GreenFormat);
				$rowcounter++;
			}
		}
		else {
			if ($z eq " " ){
			$rowcounter++;
			}
			else{
			$worksheet->write($rowcounter,$colcounter,$CodeReview_Array[$rowcounter-4], $FrozenFormat);
			$rowcounter++;
			}
		}
}
$rowcounter = 4;
$colcounter = 10;
my $index = 0;
foreach $w (@CodeReview_Array){
	#print"File Name: ".$w."\n";
	$worksheet->write($rowcounter,7, $Unique_Components[$rowcounter - 4]);
	if ($w eq " ")
	{
		$worksheet->write($rowcounter,1, $ProjectSheetID);
		$worksheet->write($rowcounter,2, $customer);
		$worksheet->write($rowcounter,3, $vehicle_platform);
		$worksheet->write($rowcounter,4, $project_id);
		$worksheet->write($rowcounter,5, $FolderLocations[$index]);
		$worksheet->write($rowcounter,6, $ForwardRels[$index]);###### or Release sheet #????
		$worksheet->write($rowcounter,8, $COSIPA_Array[$index]);
		$index++;
	}	
	$rowcounter++;
}
$rowcounter=4;
$colcounter++;

foreach $z(@QAC_RedGreenArray){
		if ($z =~ /.*\.pdf\sarchived\s(.*)/)
		{		
			if ($1 == "1.0")
			{
				$worksheet->write($rowcounter,$colcounter, $QAC_Array[$rowcounter-4], $RedFormat);
				$rowcounter++;
			}
			else
			{
				$worksheet->write($rowcounter,$colcounter,$QAC_Array[$rowcounter-4], $GreenFormat);
				$rowcounter++;
			}
		}
		else {
			if ($z eq " " ){
			$rowcounter++;
			}
			else{
			$worksheet->write($rowcounter,$colcounter,$QAC_Array[$rowcounter-4], $FrozenFormat);
			$rowcounter++;
			}
		}
}
$rowcounter = 4;
$colcounter++;

foreach $z(@CompBug_RedGreenArray){
		if ($z =~ /.*\.pdf\sarchived\s(.*)/)
		{		
			if ($1 == "1.0")
			{
				$worksheet->write($rowcounter,$colcounter, $CompBug_Array[$rowcounter-4], $RedFormat);
				$rowcounter++;
			}
			else
			{
				$worksheet->write($rowcounter,$colcounter,$CompBug_Array[$rowcounter-4], $GreenFormat);
				$rowcounter++;
			}
		}
		else {
			if ($z eq " " ){
			$rowcounter++;
			}
			else{
			$worksheet->write($rowcounter,$colcounter,$CompBug_Array[$rowcounter-4], $FrozenFormat);
			$rowcounter++;
			}
		}
}
print "\n";
print "\n";
print "Excel File has been writen to: Reports.xlsx"."\n";


