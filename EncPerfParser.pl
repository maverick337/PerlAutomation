#! c:/Perl/bin/perl -w

use Win32;
use strict;
use Spreadsheet::WriteExcel;

my ($line, @tmparr, $watermark);

open Hndl, "$ARGV[0]" or warn "Cannot open for read:$!";

# Creating new Excel file #
my $FileName = 'Report.xls';
my $workbook = Spreadsheet::WriteExcel->new($FileName);

# Adding worksheet #
my $worksheet1 = $workbook->add_worksheet('Performance');

# Defining the format and adding it to the worksheet #
my $format = $workbook->add_format(
center_across => 1,
bold => 1,
size => 10,
border => 4,
color => 'black',
bg_color => 'cyan',
border_color => 'black',
align => 'vcenter',
);

# Change width for only first column
$worksheet1->set_column(0,0,20);

$worksheet1->write(0,0, "Data & Time", $format);
$worksheet1->write(0,1, "Mem Usage", $format);
$worksheet1->write(0,2, "CPU Time", $format);
$worksheet1->write(0,3, "Total Sectors", $format);
$worksheet1->write(0,4, "Watermak", $format);

$format = $workbook->add_format();

#my $parsedata = 'parsedata.txt';
#open FILE, ">$parsedata" or warn "Cannot open for write:$!";

my ($row, $flag);
$row = 1;
$flag = 0;
while (<Hndl>) {
	$line = $_;
	
	# check for date and time #
	if ($line =~ /\w+\s+\d+\/\d+\/\d+\s+(.*)/i) {
		if ($flag == 2) {$row++;}
		#print FILE "Data & Time = $line";
		chomp($line);
		$worksheet1->write("$row",0, "$line", $format);
		$flag = 1;
	}
	
	# check for CPU time and Mem Usage #
	if ($line =~ /^pgptray\.exe\s+\d+\s+\w+\s+\d+\s+(.*)/i) {
		@tmparr = split('\s+',$line);
		#print FILE "Mem Usage = $tmparr[4] $tmparr[5] and CPU time = $tmparr[8] \n";
		$worksheet1->write("$row",1, "$tmparr[4] $tmparr[5]", $format);
		$worksheet1->write("$row",2, "$tmparr[8]", $format);
		$flag = 2;
	}
	
	# check for Watermarks #
	if ($line =~ /^\s+total\s+sectors\:\s+(.*)/i) {
		@tmparr = split('\s+',$1);
		#print FILE "Watermark = $tmparr[2] $tmparr[1] \n\n";
		$worksheet1->write("$row",3, "$tmparr[0] ", $format);
		$worksheet1->write("$row",4, "$tmparr[2] $tmparr[1]", $format);
		$flag = 3;
		$row++;
	} 
}

$workbook->close();

close Hndl;
#close FILE;

exit;