#!/usr/bin/perl
#perl2exe_include "unicore/Heavy.pl";
#perl2exe_include "DBD/ADO.pm";
#perl2exe_info ProductName=Formula_Tool
#perl2exe_info ProductVersion=2.3.5

use strict;
our ($VERSION) = 'Live_2.3.5';
print "v$VERSION\n";

#use warnings;
use FormulaTool;
use FormulaTool::LogFile;
use FormulaTool::PProRaw;
use FormulaTool::FormRaw;
use FormulaTool::FormPkg;
use FormulaTool::FormFormula;
use FormulaTool::PProFormula;
use FormulaTool::SpecFormula;
use FormulaTool::Batch;
use Date::Calc qw{Date_to_Days};

my @params = @_;
my $debug  = $params[0];
my $bell    = chr(7);
my $timer_s = time;
my @date    = localtime time;
my $day     = sprintf "%04d-%02d-%02d %02d:%02d:%02d %s", ($date[5] + 1900), ($date[4] + 1), $date[3], $date[2], $date[1], $date[0], $date[2] < 12?"AM":"PM";
my $daycode = sprintf "%04d%02d%02d", ($date[5] + 1900), ($date[4] + 1), $date[3];
my $username = uc $ENV{USERNAME};

# Global HTML
my $main_tableh = '
<div class="header">&nbsp;</div>
<table class="pagebody" cellpadding="0" cellspacing="1">
<tr style="height:15px">
    <td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td>
    <td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td>
    <td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td>
    <td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td>
    <td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td>
    <td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td>
    <td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td>
    <td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td>
    <td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td>
    <td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td><td style="width:1vw">&nbsp;</td>
 </tr>';
my $main_tablef = '
</table>
<div class="footer" style="page-break-after: always;"></div>
<hr width=100% class="noprint"/>';
my $single_blank = '<tr style="height:15px"><td colspan="100">&nbsp;</td></tr>
';
my $check_box =
  '<span style="border: black solid 1px; width:3rem">&nbsp;&nbsp;</span>';

my $wrkdir   = q{};
my $log      = LogFile->new( "FormTool-$username-$daycode.log", );

$log->log( "*** *** Start Time: " . localtime $timer_s . " $username", 2 );
$log->append( "*** Version: " . $VERSION, 2 );

my $dbcheck1 = FormulaTool->connect_formdb;
$dbcheck1->disconnect;
my $dbcheck2 = FormulaTool->connect_pprodb;
$dbcheck2->disconnect;
my $dbcheck3 = FormulaTool->connect_specdb;
$dbcheck3->disconnect;

MainMenu();

my $timer_e = time;
my $total_time = format_timediff( $timer_s, $timer_e );
$log->log( "*** *** End Time: " . localtime $timer_s . " $username", 2 );
$log->append( "*** Time running: $total_time\n\n", 2 );
exit;




sub locate_db {
#####
    #
    # locate_db(product) - find instances of product across dbs and return list
    # 			ppro::ICITEM- ITEM, ITMDESC
    #			ppro::PPBMHD- ITEM, VERSION
    #			form::FormulaMaster- FormulaCode, Description
    #			form::BOMMaster- BOMCode, Description
    #
#####
    local $| = 1;
    my @args    = @_;
    my $product = $args[0];

    $log->log( ">>>locate_db($product)", 2 );
    if ( $product !~ m/\S+/gm ) {
        $log->log( "!! Unable to process '$product'", 1 );
        return 0;
    }

    my $ppro = FormulaTool->connect_pprodb;
    print "\nProcess Pro - Item Master Table\n";
    my @result = FormulaTool->querydb( $ppro, 'ITEM, ITMDESC', 'ICITEM', "ITEM like '$product%'" );
    my $temp = Data::Dumper->Dump(@result);
    $temp =~ s/[\[\]]\n//gm;
    print "$temp";
    print "\nProcess Pro - BOM Headers Table\n";
    @result = FormulaTool->querydb( $ppro, 'ITEM,VERSION', 'PPBMHD', "ITEM like '$product%'" );
    $temp = Data::Dumper->Dump(@result);
    $temp =~ s/[\[\]]\n//gm;
    print "$temp";
    $ppro->disconnect;  
    
    my $form = FormulaTool->connect_formdb;    
    print "\nFormulator - Raw Materials Table\n";
    @result = FormulaTool->querydb( $form, 'ItemCode,Description', 'RawMaterials', "ItemCode like '$product%'" );
    $temp = Data::Dumper->Dump(@result);
    $temp =~ s/[\[\]]\n//gm;
    print "$temp";
    print "\nFormulator - Formula Master Table\n";
    @result = FormulaTool->querydb( $form, 'FormulaCode,Description', 'FormulaMaster', "FormulaCode like '$product%'" );
    $temp = Data::Dumper->Dump(@result);
    $temp =~ s/[\[\]]\n//gm;
    print "$temp";
    print "\nFormulator - BOM Master Table\n";
    @result = FormulaTool->querydb( $form, 'BOMCode,Description', 'BOMMaster', "BOMCode like '$product%'" );
    $temp = Data::Dumper->Dump(@result);
    $temp =~ s/[\[\]]\n//gm;
    print "$temp";
    $form->disconnect;

    my $spec = FormulaTool->connect_specdb;    
    print "\nSpecifications - Spec Table\n";
    @result = FormulaTool->querydb( $spec, '[Formula ID],[Customer Name],[Product Name]', '[FP Specs]', "[Formula ID] like '$product%'" );
    $temp = Data::Dumper->Dump(@result);
    $temp =~ s/[\[\]]\n//gm;
    print "$temp";
    print "\nSpecifications - SKU Table\n";
    @result = FormulaTool->querydb( $spec, '[Formula Code],[SKU Code]', '[FP Sku Index]', "[Formula Code] like '$product%'" );
    $temp = Data::Dumper->Dump(@result);
    $temp =~ s/[\[\]]\n//gm;
    print "$temp";
    $spec->disconnect;

    return 1;
}

sub dump_obj {
#####
    #
    # dump_obj(product) - Dump Objects from each DB
    # 			ppro::Raw Material
    #			ppro::Formula
    #			form::Raw Material
    #			form::Formula
    #           spec::Formula
    #
#####
    local $| = 1;
    my @args    = @_;
    my $product = $args[0];
    my $output;

    $log->log( ">>>dump_obj($product)", 2 );
    if ( $product !~ m/\S+/gm ) {
        $log->log( "!! Unable to process '$product'", 1 );
        return 0;
    }

    my $ppro = FormulaTool->connect_pprodb;
    print "\nProcess Pro - Raw Material\n";
    my $validate = FormulaTool->validate_raw($product);
    if ( $validate == 0 || $validate == 2 ) { # process if valid in PPro or Both
        my @result = PProRaw->new($product);
        $output .= "\nProcess Pro - Raw Material\n" . FormulaTool->vardump(@result);   
    } else {
        $output .= "\nProcess Pro - Raw Material\n" . "invalid code: $validate\n";
    }
    print "\nProcess Pro - Formula\n";
    $validate = FormulaTool->validate_formula($product);
    if ( $validate == 0 || $validate == 2 || $validate == 4 || $validate == 6 ) { # process if valid in PPro or All
        my @result = PProFormula->new($product);
        $output .= "\nProcess Pro - Formula\n" . FormulaTool->vardump(@result);
    } else {
        $output .= "\nProcess Pro - Formula\n" . "invalid code: $validate\n";
    }
    $ppro->disconnect;  
    
    my $form = FormulaTool->connect_formdb;    
    print "\nFormulator - Raw Material\n";
    $validate = FormulaTool->validate_raw($product);
    if ( $validate == 0 || $validate == 1 ) { # process if valid in Formulator or Both
        my @result = FormRaw->new($product);
        $output .= "\nFormulator - Raw Material\n" . FormulaTool->vardump(@result);
    } else {
        $output .= "\nFormualtor - Raw Material\n" . "invalid code: $validate\n";
    }
    print "\nFormulator - Formula\n";
    $validate = FormulaTool->validate_formula($product);
    if ( $validate == 0 || $validate == 1 || $validate == 4 || $validate == 5 ) { # process if valid in Formulator or All
        my @result = FormFormula->new($product);
        $output .= "\nFormulator - Formula\n" . FormulaTool->vardump(@result);
    } else {
        $output .= "\nFormulator - Formula\n" . "invalid code: $validate\n";
    }
    $form->disconnect;

    my $spec = FormulaTool->connect_specdb;    
    print "\nSpecifications - Formula\n";
    $validate = FormulaTool->validate_formula($product);
    if ( $validate == 0 || $validate == 1 || $validate == 2 || $validate == 3 ) { # process if valid in Specifcations or All
        my @result = SpecFormula->new($product);
        $output .= "\nSpecifications - Formula\n" . FormulaTool->vardump(@result);       
    } else {
        $output .= "\nSpecifications - Formula\n" . "invalid code: $validate\n";
    }
    $spec->disconnect;

    my $path = "$ENV{USERPROFILE}\\Desktop\\";
    my $file = sprintf "%s Dump.txt", $product;
    $file =~ s|\#||g;
    $file =~ s|/|_|g;
    $file =~ s|^a-zA-Z\d\s[()+,-]|_|g;
    my $txt = LogFile->new( $file, $path );
    $txt->write( "$output", 2 );
    $log->log( ">> $txt->{'FilePath'} completed!", 1 );

    return 1;
}

sub format_timediff {
#####
    #
    # formatTimeDiff(start time, end time) -
    #   Returns formatted text difference in times
    #
#####
    my @args  = @_;
    my $start = $args[0];
    my $end   = $args[1];
    my $diff  = $end - $start;
    my ( $hours, $mins, $sec ) = ( 0, 0, 0 );

    if ( $diff > 3600 ) {
        $hours = int( $diff / 3600 );
        $diff  = $diff % 3600;
    }
    if ( $diff > 60 ) {
        $mins = int( $diff / 60 );
        $diff = $diff % 60;
    }
    $sec = int $diff;

    return "$hours h, $mins m, $sec s";
}

sub xfer_item2p {
#####
    #
    # xfer_item2p(product,flag) - Transfer Item from Formulator to ProcessPro
    #
#####
    local $| = 1;
    my @args    = @_;
    chomp $args[0];
    my $product = $args[0];
    my $flag    = $args[1];

    my ($statement,$statement2);
    my $validate = FormulaTool->validate_raw($product);
    if ( $flag == 2 || $validate == 2 || $validate == 3 ) { return 0; } # Return if invalid in Formulator or Both
    $log->log( " >>xfer_item2ppro $product", 2 );

    #####

    print "\n*** Transfering Raw Material from Formulator to Process Pro\n\n";

    my $form_item = FormRaw->new($product);
    my $ppro_item = PProRaw->new($product);
    my $new_item = PProRaw->new($product);

    $form_item->{'Description'} =~ s/\s\s+//m;
    $form_item->{'Description'} =~ s/\'//m;
    $form_item->{'Class'} =~ s/\s\s+//m;
    $form_item->{'SubClass'} =~ s/\s\s+//m;
    $form_item->{'Vendor'} =~ s/\s\s+//m;

    $ppro_item->{'Description'} = substr $ppro_item->{'Description'}, 0, 59;
    $ppro_item->{'Description'} =~ s/\s\s+//m;

    #####

    $new_item->{'Description'} = substr $form_item->{'Description'}, 0, 59;
    if (!$new_item->{'Class'}) { $new_item->{'Class'} = $form_item->{'Class'}; }
    if (!$new_item->{'SubClass'}) { $new_item->{'SubClass'} = $form_item->{'SubClass'}; }
    if (!$new_item->{'VendorCode'}) { $new_item->{'VendorCode'} = $form_item->{'VendorCode'}; }

    $new_item->{'Status'} = 'A';
    $new_item->{'StdCost'} = $form_item->{'Cost'};
    if (!$new_item->{'PurchUOM'}) { $new_item->{'PurchUOM'} = uc $form_item->{'CostUOMDescr'}; }
    if ($new_item->{'PurchUOM'} =~ m/EACH/ig) {
        $new_item->{'SalesUOM'} =  $new_item->{'StockUOM'} = $new_item->{'CountUOM'} = $new_item->{'CompUOM'} = 'EACH';
    } else {
        $new_item->{'SalesUOM'} =  $new_item->{'StockUOM'} = $new_item->{'CountUOM'} = $new_item->{'CompUOM'} = 'KG';
    }
    
    if ( $product =~ m/^(([A-z]|99)\.)?\d\d\d\d\d\d$/ ) {
        $new_item->{'MkInternal?'}  = 0;
        $new_item->{'PurExternal?'} = 1;
        $new_item->{'Class'}        = 'RM';
        $new_item->{'Resell?'}      = 1;
        $new_item->{'GLAccounts'}[0] = '12000-000-0000';
        $new_item->{'GLAccounts'}[1] = '20100-000-0000';
        $new_item->{'GLAccounts'}[2] = '58000-000-0000';
        $new_item->{'GLAccounts'}[3] = '26100-000-0000';
        $new_item->{'GLAccounts'}[4] = '57400-000-0000';
        $new_item->{'GLAccounts'}[5] = '57450-000-0000';
        $new_item->{'GLAccounts'}[6] = '12090-000-0000';
        $new_item->{'GLAccounts'}[7] = '26100-000-0000';
        $new_item->{'GLAccounts'}[8] = '57400-000-0000';
        $new_item->{'GLAccounts'}[9] = '57450-000-0000';
        $new_item->{'GLAccounts'}[10] = '59990-000-0000';
        $new_item->{'GLAccounts'}[11] = '59990-000-0000';          
    } else {
        $new_item->{'MkInternal?'}  = 1;
        $new_item->{'PurExternal?'} = 0;
        $new_item->{'Class'}        = 'BL';
        $new_item->{'SubClass'}     = 'BLEND';
        $new_item->{'Resell?'}      = 0;
        $new_item->{'PurchUOM'}     = 'KG';
        $new_item->{'GLAccounts'}[0] = '12100-000-0000';
        $new_item->{'GLAccounts'}[1] = '20100-000-0000';
        $new_item->{'GLAccounts'}[2] = '58100-000-0000';
        $new_item->{'GLAccounts'}[3] = '26100-000-0000';
        $new_item->{'GLAccounts'}[4] = '57400-000-0000';
        $new_item->{'GLAccounts'}[5] = '57450-000-0000';
        $new_item->{'GLAccounts'}[6] = '12110-000-0000';
        $new_item->{'GLAccounts'}[7] = '26100-000-0000';
        $new_item->{'GLAccounts'}[8] = '57400-000-0000';
        $new_item->{'GLAccounts'}[9] = '57450-000-0000';
        $new_item->{'GLAccounts'}[10] = '59990-000-0000';
        $new_item->{'GLAccounts'}[11] = '59990-000-0000';
    }

    if ( $validate == 1 ) {
        if ( !$flag ) {
            print "\n  Adding to ProcessPro :\t$form_item->{'ItemCode'} [$form_item->{'Description'}]\n\n";
            my $input = prompt("\tDo you wish to continue [default: Y] ? ");
            if ( ( $input !~ m/^y/i ) && ($input) ) {
                $log->log( ">> Transfer canceled.", 1 );
                return 1;
            }
        }
        $statement = sprintf "INSERT INTO ICITEM (ITEM, ITMDESC, PLINID, STKUMID, SUNMSID, PUNMSID, IONHAND, DECNUM, STKCODE, TYPE, RESELL, USELOTS, USESERL, HISTORY, MAKEITR, ITMCLSS, CODE, ITEMSTAT, NOMONTHS, MTHSORDAYS, CUNMSID, CMPUMID, ADDUSER, ADDDATE, PUREXTR, STDCOST) VALUES ('%s', '%s', 'RAWMAT', '%s', '%s', '%s', 0, 3, 'Y', 'I', 1, 'S', 'N', 1, %d, '%s', '%s', '%s', 24, 1, '%s', '%s', '%s', '%s', %d, %0.2f)", $new_item->{'ItemCode'}, $new_item->{'Description'}, $new_item->{'StockUOM'}, $new_item->{'SalesUOM'}, $new_item->{'PurchUOM'}, $new_item->{'MkInternal?'}, $new_item->{'Class'}, $new_item->{'SubClass'}, $new_item->{'Status'}, $new_item->{'CountUOM'}, $new_item->{'CompUOM'}, $username, $day, $new_item->{'PurExternal?'}, $new_item->{'StdCost'};
    } else {
        if ( !$flag ) {
            print "\n  Formulator :\t$form_item->{'ItemCode'} [$form_item->{'Description'}]\n  ProcessPro :\t$ppro_item->{'ItemCode'} [$ppro_item->{'Description'}]\n\n";
            my $input = prompt("\tDo you wish to update Process Pro [default: Y] ? ");
            if ( ( $input !~ m/^y/i ) && ($input) ) {
                $log->log( ">> Transfer canceled.", 1 );
                return 1;
            }
        }
        $statement = sprintf "UPDATE ICITEM SET ITMDESC='%s', PLINID='RAWMAT', STKUMID= '%s', SUNMSID='%s', PUNMSID='%s', DECNUM=3, STKCODE='Y', TYPE='I', RESELL=1, USELOTS='S', USESERL='N', HISTORY=1, MAKEITR=%d, ITMCLSS='%s', CODE='%s', ITEMSTAT='%s', NOMONTHS=24, MTHSORDAYS=1, CUNMSID='%s', LCKUSER='%s', LCKDATE='%s', PUREXTR=%d, CMPUMID='%s', STDCOST=%.02f WHERE ITEM='%s'", $new_item->{'Description'}, $new_item->{'StockUOM'}, $new_item->{'SalesUOM'}, $new_item->{'PurchUOM'}, $new_item->{'MkInternal?'}, $new_item->{'Class'}, $new_item->{'SubClass'}, $new_item->{'Status'}, $new_item->{'CountUOM'}, $username, $day, $new_item->{'PurExternal?'}, $new_item->{'CompUOM'}, $new_item->{'StdCost'}, $new_item->{'ItemCode'};
    }
    $log->log( " > $statement", 2 );
    
    my $dbh = FormulaTool->connect_pprodb;
    my $sth = $dbh->do($statement)
      or $log->log("!! Could not do [$statement] \nerror: $!");
    if (($ppro_item->{'AllergenCheck'}>0)&&($form_item->{'AllergenCheck'}>0)) {
        $statement2 = sprintf "UPDATE PPALGN SET WHEAT=%d, MILK=%d, SOY=%d, EGGS=%d, TREENUTS=%d, PEANUTS=%d, FISH=%d, SEAFOOD=%d, LCKUSER='%s', LCKDATE='%s' WHERE ITEM='%s';",$form_item->{'Allergens'}[0][0], $form_item->{'Allergens'}[1][0], $form_item->{'Allergens'}[2][0], $form_item->{'Allergens'}[3][0], $form_item->{'Allergens'}[4][0], $form_item->{'Allergens'}[5][0], $form_item->{'Allergens'}[6][0], $form_item->{'Allergens'}[7][0], $username, $day, $new_item->{'ItemCode'};
        $log->log( " > $statement2", 2 );   
        $sth = $dbh->do($statement2) or $log->log("!! Could not do [$statement2] \nerror: $!");
    } elsif ($form_item->{'AllergenCheck'}>0) {
       $statement2 = sprintf "INSERT INTO PPALGN (ITEM, WHEAT, MILK, SOY, EGGS, TREENUTS, PEANUTS, FISH, SEAFOOD, ADDUSER, ADDDATE) VALUES ('%s', %d, %d, %d, %d, %d, %d, %d, %d, '%s', '%s')", $new_item->{'ItemCode'}, $form_item->{'Allergens'}[0][0], $form_item->{'Allergens'}[1][0], $form_item->{'Allergens'}[2][0], $form_item->{'Allergens'}[3][0], $form_item->{'Allergens'}[4][0], $form_item->{'Allergens'}[5][0], $form_item->{'Allergens'}[6][0], $form_item->{'Allergens'}[7][0], $username, $day;
        $log->log( " > $statement2", 2 );   
        $sth = $dbh->do($statement2) or $log->log("!! Could not do [$statement2] \nerror: $!");
    }
    $dbh->disconnect;

    #####

    my $bell = chr(7);
    if ( !$new_item->{'WH1?'} ) {
        $statement = sprintf "INSERT INTO ICILOC (LOCTID, ITEM, ICACCT, RCLACCT, ICLACCT, DLBACCT, FXOACCT, VROACCT, MIPACCT, DIPACCT, FIPACCT, VIPACCT, MCLACCT, ICSHPCLEAR, ADDUSER, ADDDATE, QCSTAT, BUYER) VALUES ('WH1', '%s', '%s', '%s', '%s','%s', '%s','%s', '%s','%s', '%s','%s', '%s', '%s', '%s', '%s', 'Q', 'PM')", $new_item->{'ItemCode'}, $new_item->{'GLAccounts'}[0], $new_item->{'GLAccounts'}[1], $new_item->{'GLAccounts'}[2], $new_item->{'GLAccounts'}[3], $new_item->{'GLAccounts'}[4], $new_item->{'GLAccounts'}[5], $new_item->{'GLAccounts'}[6], $new_item->{'GLAccounts'}[7], $new_item->{'GLAccounts'}[8], $new_item->{'GLAccounts'}[9], $new_item->{'GLAccounts'}[10], $new_item->{'GLAccounts'}[11],$username,$day;
        printf "$bell\n*** You will need to setup a WHS for %s [%s]!***\t$bell\n",$new_item->{'ItemCode'},$new_item->{'Description'};
    } else {
        $statement = sprintf "UPDATE ICILOC SET ICACCT='%s', RCLACCT='%s', ICLACCT='%s', DLBACCT='%s', FXOACCT='%s', VROACCT='%s', MIPACCT='%s', DIPACCT='%s', FIPACCT='%s', VIPACCT='%s', MCLACCT='%s', ICSHPCLEAR='%s', LCKUSER='%s', LCKDATE='%s', QCSTAT='Q', BUYER='PM' WHERE ITEM='%s'", $new_item->{'GLAccounts'}[0], $new_item->{'GLAccounts'}[1], $new_item->{'GLAccounts'}[2], $new_item->{'GLAccounts'}[3], $new_item->{'GLAccounts'}[4], $new_item->{'GLAccounts'}[5], $new_item->{'GLAccounts'}[6], $new_item->{'GLAccounts'}[7], $new_item->{'GLAccounts'}[8], $new_item->{'GLAccounts'}[9], $new_item->{'GLAccounts'}[10], $new_item->{'GLAccounts'}[11], $username, $day, $new_item->{'ItemCode'};
    }

    $log->log( " > $statement", 2 );
    $dbh = FormulaTool->connect_pprodb;
    $sth = $dbh->do($statement)
      or $log->log("!! Could not do [$statement] \nerror: $!");
    $dbh->disconnect;
    $log->log(">> Formulator Raw Material $new_item->{'ItemCode'} [$new_item->{'Description'}] transferred to ProcessPro!", 1 );

    return 1;
}

sub xfer_form2p {
#####
    #
    # xfer_form2p - Transfer formula from Formulator to ProcessPro
    #
#####
    local $| = 1;
    my @args    = @_;
    chomp $args[0];
    my $product = $args[0];

    my $statement;
    my $reason;
    my $validate = FormulaTool->validate_formula($product);
    if ( $validate == 2 || $validate == 3 || $validate == 6 || $validate == 7) { return 0; } # Return if invalid in Formulator or Both

    $log->log( ">>>xfer_form2ppro $product", 2 );
    my $ppro_form = PProFormula->new($product);
    my $form_form = FormFormula->new($product);
    my $spec_form = SpecFormula->new($product);

    if (!$form_form->{'RawMaterials'}[0][7]) {
        $log->log("!! No percentages found in Formulator:\n\tPlease resave the formula in Formulator, then retry the transfer.",1); return 0;
    }

    my $c1 = sprintf "%.4f",$form_form->{'FormulaTotal'} * $form_form->{'ServingQty'};
    my $c2 = sprintf "%.4f",$form_form->{'ServingTotal'};
    
    if ($c1 != $c2) {
    $log->log(" > Formula Total: $form_form->{'FormulaTotal'} @ $form_form->{'ServingQty'} units = $c1",1);
    $log->log(" > Serving Total: $form_form->{'ServingTotal'}",1);
    $log->log("!! Formula Total weight does not match Serving Size:\n\tPlease adjust the Serving Size in Formulator to $c1\n\tthen retry the transfer.",1); return 0;
    }
    print "\n";

    if ($ppro_form->{'Revision'} =~ m/^\d+$/i) { $ppro_form->{'Revision'}++; } else { $ppro_form->{'Revision'} = 1; }

    if (!$ppro_form->{'RevDate'} =~ m/1900-01-01/i) {
        $ppro_form->{'RevDate'} =~ m/(\d\d\d\d)-(\d\d)-(\d\d)./i;
        my $pprodays = Date_to_Days($1,$2,$3);
        my $todays = Date_to_Days(($date[5] + 1900), ($date[4] + 1), $date[3]);
        if ($pprodays < $todays) {
            if (($date[4]+1 == 2)&&($date[3] == 29)) {
                $ppro_form->{'RevDate'} = ($date[5] + 1900).'-03-01';  
            } else {
                $ppro_form->{'RevDate'} = ($date[5] + 1900).'-'.($date[4] + 1).'-'.$date[3];
            }
        }
    } else {
        if (($date[4]+1 == 2)&&($date[3] == 29)) {
            $ppro_form->{'RevDate'} = ($date[5] + 1900).'-03-01';  
        } else {
            $ppro_form->{'RevDate'} = ($date[5] + 1900).'-'.($date[4] + 1).'-'.$date[3];
        }
    }

    while (1) {
        print "Revision: " . $ppro_form->{'Revision'} . "\n";
        print "Rev Date: " . $ppro_form->{'RevDate'}. "\n";
        print "\tEnter the reason for the revision:  ";
        $reason = <>;
        chomp $reason;
        if ((!$reason)||(length $reason<2)) {
            next;
        } else {
            last;
        }
    }

    #####

    my $cs=0;
    foreach my $count ( 0 .. $#{ $form_form->{'RawMaterials'} } ) {
        if ( $form_form->{'RawMaterials'}[$count][2] =~ m/2\d{5}/i ) { $cs++; }
    }
    $log->log(" > Found $cs Customer Supplied ingredients",2);
    if ( ( $form_form->{'ServingDesc'} =~ m/Pow/ig ) || ( $cs > 0 ) ) {
        $ppro_form->{'Yield'} = 100;
    } else { $ppro_form->{'Yield'} = 97; }

    if ( $validate == 1 || $validate == 5 ) {
        print "\n  Adding to ProcessPro :\t$form_form->{'FormulaCode'} [$form_form->{'Description'}]\n\n";
        my $input = prompt("\tDo you wish to continue [default: Y] ? ");
        if ( ( $input !~ m/^y/i ) && ($input) ) {
            $log->log( ">> Transfer canceled.", 1 );
            return 1;
        }
        my $complete = xfer_item2p( $product, 1 );
        if ( !$complete ) { return 0; }
        $statement = sprintf "INSERT INTO PPBMHD (ITEM, VERSION, TYPE, ROUTE, PARTIAL, BATCHES, YIELD, OVERHEAD, REVISION, REVSTAT, INACTIVE, ADDUSER, ADDDATE, REVDATE, NOTES) VALUES ('%s', '%s', 1, 'BLENDING', 1, 1, '%s', 'PWDLB1', '%s', 'N', 0, '%s', '%s', '%s', '%s')", $form_form->{'FormulaCode'}, $ppro_form->{'Version'}, $ppro_form->{'Yield'}, $ppro_form->{'FormulaCode'}, $username, $day, $ppro_form->{'RevDate'}, $form_form->{'Notes'};
    } else {
        print "\n  Formulator :\t$form_form->{'FormulaCode'} [$form_form->{'Description'}]\n  ProcessPro :\t$ppro_form->{'FormulaCode'} [$ppro_form->{'Description'}]\n\n";
        my $input = prompt("\tDo you wish to update Process Pro [default: Y] ? ");
        if ( ( $input !~ m/^y/i ) && ($input) ) {
            $log->log( ">> Transfer canceled.", 1 );
            return 1;
        }
        my $complete = xfer_item2p( $product, 1 );
        if ( !$complete ) { return 0; }
        $statement =  sprintf "UPDATE PPBMHD SET TYPE=1, ROUTE='BLENDING', PARTIAL=1, BATCHES=1, YIELD=%d, OVERHEAD='PWDLB1', REVISION='%s', REVSTAT='N', INACTIVE=0, LCKDATE='%s', LCKUSER='%s', REVDATE='%s', NOTES='%s' WHERE ITEM='%s' AND VERSION like '%s'", $ppro_form->{'Yield'}, $ppro_form->{'Revision'}, $day, $username, $ppro_form->{'RevDate'}, $form_form->{'Notes'}, $ppro_form->{'FormulaCode'}, $ppro_form->{'Version'};
    }
    $log->log( " > $statement", 2 );
    my $dbh = FormulaTool->connect_pprodb;
    my $sth = $dbh->do($statement)
      or $log->log("!! Could not do [$statement] \nerror: $!");

    #####

    $log->log(" > Dropping BOM Details for $form_form->{'FormulaCode'} [$form_form->{'Description'}] Version $ppro_form->{'Version'}...", 1 );
    $statement = "DELETE FROM PPBMDT WHERE ITEM like '$form_form->{'FormulaCode'}' AND VERSION like '$ppro_form->{'Version'}'";
    $log->log(" > $statement", 2);
    $sth = $dbh->do($statement)
      or $log->log("!! Could not do [$statement] \nerror: $!");

    foreach my $count ( 0 .. $#{ $form_form->{'RawMaterials'} } ) {
        if ( $form_form->{'RawMaterials'}[$count][1] != 1 ) { next; }
        $log->log(" > Adding $form_form->{'RawMaterials'}[$count][2] to BOM details...", 2 );
        $statement = sprintf "INSERT INTO PPBMDT (ITEM, PARTNO, FINDNUM, QTY, VERSION, ADDUSER, ADDDATE) VALUES ('%s', '%s', %d, %.07f, '%s', '%s', '%s')", $form_form->{'FormulaCode'}, $form_form->{'RawMaterials'}[$count][2], $form_form->{'RawMaterials'}[$count][0] * 10, $form_form->{'RawMaterials'}[$count][7] / 100, $ppro_form->{'Version'}, $username, $day;
        $log->log(" > $statement", 2);
        $sth = $dbh->do($statement)
          or $log->log( "!! Could not do [$statement] \nerror: $!", 1 );
    }

    $statement = sprintf "INSERT INTO PPPREV (ITEM, VERSION, REVISION, ADDUSER, ADDDATE, REVMEMO) VALUES('%s', '01', '%s', '%s', '%s', '%s')", $form_form->{'FormulaCode'}, $ppro_form->{'Revision'}, $username, $day, "Transfered from Formulator by $username on $day\nReason: $reason";
    $log->log( " > $statement", 2 );
    $sth = $dbh->do($statement)
      or $log->log("!! Could not do [$statement] \nerror: $!");

    my ( $cap_add, $cap_adj, $note );
    
    if ( $form_form->{'ServingSize'} =~ m/^Size 3\w*/ ) { $cap_add = 50; $cap_adj = 3; }
    if ( $form_form->{'ServingSize'} =~ m/^Size 2\w*/ ) { $cap_add = 60; $cap_adj = 4; }
    if ( $form_form->{'ServingSize'} =~ m/^Size 1\w*/ ) { $cap_add = 75; $cap_adj = 5; }
    if ( $form_form->{'ServingSize'} =~ m/^Size 0\w*/ ) { $cap_add = 100; $cap_adj = 6; }
    if ( $form_form->{'ServingSize'} =~ m/^Size 0el\w*/ ) { $cap_add = 110; $cap_adj = 7; }
    if ( $form_form->{'ServingSize'} =~ m/^Size 00\w*/ ) {$cap_add = 120; $cap_adj = 7; }
    if ( $form_form->{'ServingSize'} =~ m/^Size 00el\w*/ ) { $cap_add = 130; $cap_adj = 10; }
    if ( $form_form->{'ServingSize'} =~ m/^Size 000\w*/ ) { $cap_add = 165; $cap_adj = 10; }   

    my $fillweight = $form_form->{'FormulaTotal'};
    my $lo_weight  = $fillweight + $cap_add;
    my $hi_weight  = ($fillweight + $cap_add ) * 1.05;
    my $unittype;
       
    if ( $form_form->{'ServingDesc'} =~ m/cap/i ) {
        $note = sprintf "Encapsulate at %d mg to %d mg", $lo_weight, $hi_weight;
        $unittype = 'CAPSUL';
    }
    elsif ( $form_form->{'ServingDesc'} =~ m/tab/i ) {
        $note = sprintf "Tablet 10ct Avg at %d mg to %d mg\n\nHardness Target: %0.1f Kp\nHardness Range: %s\n\nThickness Target: %0.2f mm\nThickness Range: %s\n", $fillweight, $fillweight * 1.05, $spec_form->{'UnitHardnessSpec'}?$spec_form->{'UnitHardnessSpec'}:($form_form->{'Hardness'}?$form_form->{'Hardness'}:"TBD"),$spec_form->{'UnitHardnessRange'}?$spec_form->{'UnitHardnessRange'}:($form_form->{'Hardness'}?$form_form->{'Hardness'}*0.4." - ".$form_form->{'Hardness'}*1.6." Kp":"TBD"),$spec_form->{'UnitThicknessSpec'}?$spec_form->{'UnitThicknessSpec'}:($form_form->{'Thickness'}?$form_form->{'Thickness'}:"TBD"),$spec_form->{'UnitThicknessRange'}?$spec_form->{'UnitThicknessRange'}:($form_form->{'Thickness'}?$form_form->{'Thickness'}*0.95." - ".$form_form->{'Thickness'}*1.05." mm":"TBD");
        $unittype = 'TABLET';
    }

    $log->log(" > Serving Desc: $form_form->{'ServingQty'} $form_form->{'ServingSize'} $form_form->{'ServingDesc'}",1);
    $log->log(" > Fill weight: $fillweight",1);
    $log->log(" > $note",1);

    my $valid = FormulaTool->validate_formula($form_form->{'FormulaCode'} . '.PR');
    if ( defined $ppro_form->{'PRWeight'} && ($valid == 0 || $valid == 2 || $valid == 4 || $valid == 6)) {
        $statement = sprintf "UPDATE PPBMDT SET QTY='%.07f', LCKUSER='%s', LCKDATE='%s' WHERE ITEM like '%s.PR' AND PARTNO like '%s' AND VERSION = '%s';", $form_form->{'FormulaTotal'} / 1000, $username, $day,$form_form->{'FormulaCode'}, $form_form->{'FormulaCode'}, $ppro_form->{'Version'};
        $log->log( " > $statement", 2 );
        $sth = $dbh->do($statement)
          or $log->log( "!! Could not do [$statement] \nerror: $!", 1 );
        $statement = sprintf "UPDATE PPBMHD SET NOTES='%s', REVISION='%s', REVSTAT='N', INACTIVE=0, REVDATE='%s' WHERE ITEM like '%s.PR' AND VERSION = '%s';", $form_form->{'PRNotes'}."\n".$note, $ppro_form->{'Revision'}, $ppro_form->{'RevDate'}, $form_form->{'FormulaCode'}, $ppro_form->{'Version'};
        $log->log( " > $statement", 2 );
        $sth = $dbh->do($statement)
          or $log->log( "!! Could not do [$statement] \nerror: $!", 1 );
    } elsif (( $form_form->{'ServingDesc'} =~ m/cap/i ) || ( $form_form->{'ServingDesc'} =~ m/tab/i ))  {
        my @results = FormulaTool->querydb( $dbh, "ITEM", "ICITEM", "ITEM like '$form_form->{\"FormulaCode\"}.PR'" );
        if ( !defined $results[0][0][0] ) {
            $statement =  sprintf "INSERT INTO ICITEM (ITEM, ITMDESC, PLINID, STKUMID, SUNMSID, PUNMSID, IONHAND, DECNUM, STKCODE, TYPE, RESELL, USELOTS, USESERL, HISTORY, MAKEITR, ITMCLSS, CODE, ITEMSTAT, NOMONTHS, MTHSORDAYS, CUNMSID, ADDUSER, ADDDATE, PUREXTR) VALUES ('%s.PR', '%s', 'RAWMAT', 'M', 'M', 'M', 0, 3, 'Y', 'I', 0, 'S', 'N', 1, 1, 'PR', '%s', 'A', 24, 1, 'M', '%s', '%s', 0)", $form_form->{'FormulaCode'}, $form_form->{'Description'}, $unittype, $username, $day;
            $log->log( " > $statement", 2 );
            $sth = $dbh->do($statement)
              or $log->log("!! Could not do [$statement] \nerror: $!");
            $statement =  sprintf "INSERT INTO ICILOC (LOCTID, ITEM, ICACCT, RCLACCT, ICLACCT, DLBACCT, FXOACCT, VROACCT, MIPACCT, DIPACCT, FIPACCT, VIPACCT, MCLACCT, ICSHPCLEAR, ADDUSER, ADDDATE, QCSTAT) VALUES ('WH1', '%s.PR', '%s', '%s', '%s','%s', '%s','%s', '%s','%s', '%s','%s', '%s', '%s', '%s', '%s', 'Q')", $form_form->{'FormulaCode'}, '12120-000-0000', '20100-000-0000', '58120-000-0000', '26100-000-0000', '57400-000-0000', '57450-000-0000', '12110-000-0000', '26100-000-0000', '57400-000-0000', '57450-000-0000', '59990-000-0000', '59990-000-0000',$username, $day;
            $log->log( " > $statement", 2 );
            $sth = $dbh->do($statement)
              or $log->log("!! Could not do [$statement] \nerror: $!");
            $statement = sprintf "INSERT INTO PPPREV (ITEM, VERSION, REVISION, ADDUSER, ADDDATE, REVMEMO) VALUES('%s.PR', '01', '%s', '%s', '%s', '%s')", $form_form->{'FormulaCode'}, $ppro_form->{'Revision'}, $username, $day, "Transfered from Formulator by $username on $day\nReason: $reason";
            $log->log( " > $statement", 2 );
            $sth = $dbh->do($statement)
              or $log->log("!! Could not do [$statement] \nerror: $!");
        }
        @results = FormulaTool->querydb( $dbh, "ITEM", "PPBMHD", "ITEM like '$form_form->{\"FormulaCode\"}.PR'" );
        if ( !defined $results[0][0][0] ) {
            my ( $route, $overhead );
            if ( $form_form->{'ServingDesc'} =~ m/cap/i ) {
                $route    = 'ENCAPULATING';
                $overhead = 'CAPLB1';
            } else {
                $route    = 'TABLETING';
                $overhead = 'TABLB1';
            }
            $statement = sprintf "INSERT INTO PPBMHD (ITEM, VERSION, TYPE, ROUTE, PARTIAL, BATCHES, YIELD, OVERHEAD, REVISION, REVSTAT, INACTIVE, ADDUSER, ADDDATE, NOTES, REVDATE) VALUES('%s.PR', '01', 1, '%s', 1, 1, 100, '%s', '%s.PR', 'N', 0, '%s', '%s', '%s', '%s')", $form_form->{'FormulaCode'}, $route, $overhead, $form_form->{'FormulaCode'}, $username, $day, $form_form->{'PRNotes'}."\n".$note, $ppro_form->{'RevDate'};
            $log->log( " > $statement", 2 );
            $sth = $dbh->do($statement)
              or $log->log("!! Could not do [$statement] \nerror: $!");
        }
        $statement = sprintf "INSERT INTO PPBMDT (ITEM, PARTNO, FINDNUM, QTY, VERSION, ADDUSER, ADDDATE) VALUES('%s.PR', '%s', 10, %.07f, '01', '%s', '%s')", $form_form->{'FormulaCode'}, $form_form->{'FormulaCode'}, $form_form->{'FormulaTotal'} / 1000, $username, $day;
        $log->log( " > $statement", 2 );
        $sth = $dbh->do($statement)
              or $log->log("!! Could not do [$statement] \nerror: $!");
        $statement = sprintf "INSERT INTO PPPREV (ITEM, VERSION, REVISION, ADDUSER, ADDDATE, REVMEMO) VALUES('%s.PR', '01', '%s', '%s', '%s', '%s')", $form_form->{'FormulaCode'}, $ppro_form->{'Revision'}, $username, $day, "Transfered from Formulator by $username on $day\nReason: $reason";
            $log->log( " > $statement", 2 );
            $sth = $dbh->do($statement)
              or $log->log("!! Could not do [$statement] \nerror: $!");
    } else {
        $log->log(" > No PR needed.", 1);
    }

    #####

    $log->log(">> Formulator Formula $form_form->{'FormulaCode'} [$form_form->{'Description'}] transferred to ProcessPro!", 1 );
    $dbh->disconnect;
    return 1;
}


sub xfer_pkg2f {
#####
    #
    # xfer_pkg2f - Transfer Packaging from ProcessPro to Formulator
    #
#####
    local $| = 1;
    my @args    = @_;
    chomp $args[0];
    my $product = $args[0];
    my $flag    = $args[1];

    my ($statement,$statement2);
    my $validate = FormulaTool->validate_pkg($product);
    if ( $flag == 2 || $validate == 1 || $validate == 3 ) { return 0; } # Return if invalid in PPro or Both
    $log->log( " >>xfer_pkg2form $product", 2 );

    #####

    print "\n*** Transfering Packaging from Process Pro to Formulator\n\n";

    my $form_item = FormPkg->new($product);
    my $ppro_item = PProRaw->new($product);
    my $new_item = FormPkg->new($product);

    $ppro_item->{'ItemCode'} =~ s/\s\s+//m;
    $ppro_item->{'Description'} =~ s/\s\s+//m;
    $ppro_item->{'VendorCode'} =~ s/\s\s+//m;
    $ppro_item->{'AvgCost'} = sprintf '%.2f', $ppro_item->{'AvgCost'};
    $ppro_item->{'StdCost'} = sprintf '%.2f', $ppro_item->{'StdCost'};

    $form_item->{'Description'} = substr $form_item->{'Description'}, 0, 59; 
    $form_item->{'Description'} =~ s/\s\s+//m;

    #####

    $new_item->{'Description'} = substr $ppro_item->{'Description'}, 0, 59;
    $new_item->{'Class'} = $ppro_item->{'Class'};
    $new_item->{'SubClass'} = $ppro_item->{'SubClass'};    
    if (!$new_item->{'VendorCode'}) { $new_item->{'VendorCode'} = $ppro_item->{'VendorCode'}; }
    if ($ppro_item->{'Status'} eq "A") { $new_item->{'Status'} = 'Active'; } else { $new_item->{'Status'} = 'Inactive'; }
    if (!$new_item->{'CostUOMDescr'}) { $new_item->{'CostUOMDescr'} = $ppro_item->{'PurchUOM'}; }
    if (!$new_item->{'PurchFactor'}) { $new_item->{'PurchFactor'} = 1; }
    $new_item->{'Cost'} = ($ppro_item->{'StdCost'} != 0?$ppro_item->{'StdCost'}:$ppro_item->{'AvgCost'});
    if ( $ppro_item->{'PurchUOM'} =~ m/^M$/gi ) {
        $new_item->{'PurchFactor'} = 1000;
    }
    $new_item->{'Cost'} = $new_item->{'Cost'}/$new_item->{'PurchFactor'};

    if ( $validate == 2 ) {
        if ( !$flag ) {
            print "\n\tAdding to Formulator :\t$ppro_item->{'ItemCode'} [$ppro_item->{'Description'}]\n\n";
            my $input = prompt("\tDo you wish to continue [default: Y] ? ");
            if ( ( $input !~ m/^y/i ) && ($input) ) { return 0; }
        }
        $statement = sprintf
"INSERT INTO BOMPackaging (InstrCode, InstrText, Description, Cost, Status, PurchasingFactor, VendorCode, AddedBy, Added, Class, SubClass) VALUES ('%s', '%s', '%s', %0.5f, 'Active', %d, '%s', '%s', '%s', '%s', '%s'); ",$new_item->{'ItemCode'},$new_item->{'Description'},$new_item->{'CostUOMDescr'},$new_item->{'Cost'},$new_item->{'PurchFactor'},$new_item->{'VendorCode'},$username,$day,$new_item->{'Class'},$new_item->{'SubClass'};
        $log->log( " > $statement", 2 );
        my $dbh = FormulaTool->connect_formdb;
        my $sth = $dbh->do($statement) or $log->log("!! Could not do [$statement] \nerror: $!");
        $dbh->disconnect;
    } else {
        if ( !$flag ) {
            print "\n\tProcess Pro :\t$ppro_item->{'ItemCode'} [$ppro_item->{'Description'}]\n\tFormulator :\t$form_item->{'ItemCode'} [$form_item->{'Description'}]\n\n";
            my $input = prompt("\tDo you wish to update Formulator [default: Y] ? ");
            if ( ( $input !~ m/^y/i ) && ($input) ) { return 0; }
        }
        $statement = sprintf "UPDATE BOMPackaging SET InstrText='%s', Description='%s', Cost=%0.5f, Status='Active', PurchasingFactor=%d, VendorCode='%s', UpdatedBy='%s', Updated='%s', Class='%s', SubClass='%s' WHERE InstrCode = '%s'",$new_item->{'Description'},$new_item->{'CostUOMDescr'},$new_item->{'Cost'},$new_item->{'PurchFactor'},$new_item->{'VendorCode'},$username,$day,$new_item->{'Class'},$new_item->{'SubClass'},$new_item->{'ItemCode'};
        $log->log( " > $statement", 2 );
        my $dbh = FormulaTool->connect_formdb;
        my $sth = $dbh->do($statement)
          or $log->log("!! Could not do [$statement] \nerror: $!");
        $dbh->disconnect;
    }

    #####

    $log->log(">> ProcessPro Packaging \t$ppro_item->{'ItemCode'} [$ppro_item->{'Description'}] transferred to Formulator!", 1 );
    return 1;
}


sub xfer_item2f {
#####
    #
    # xfer_item2f - Transfer Item from ProcessPro to Formulator
    #
#####
    local $| = 1;
    my @args    = @_;
    chomp $args[0];
    my $product = $args[0];
    my $flag    = $args[1];

    my ($statement,$statement2);
    my $validate = FormulaTool->validate_raw($product);
    if ( $flag == 2 || $validate == 1 || $validate == 3 ) { return 0; }
    $log->log( " >>xfer_item2form $product", 2 );

    #####

    print "\n*** Transfering Raw Material from Process Pro to Formulator\n\n";

    my $form_item = FormRaw->new($product);
    my $ppro_item = PProRaw->new($product);
    my $new_item = FormRaw->new($product);

    $ppro_item->{'ItemCode'} =~ s/\s\s+//m;
    $ppro_item->{'Description'} =~ s/\s\s+//m;
    $ppro_item->{'Class'} =~ s/\s\s+//m;
    $ppro_item->{'SubClass'} =~ s/\s\s+//m;
    $ppro_item->{'VendorCode'} =~ s/\s\s+//m;
    $ppro_item->{'AvgCost'} = sprintf '%.2f', $ppro_item->{'AvgCost'};
    $ppro_item->{'StdCost'} = sprintf '%.2f', $ppro_item->{'StdCost'};

    $form_item->{'Description'} = substr $form_item->{'Description'}, 0, 59; 
    $form_item->{'Description'} =~ s/\s\s+//m;

    #####

    $new_item->{'Description'} = substr $ppro_item->{'Description'}, 0, 59;
    if (!$new_item->{'Class'}) { $new_item->{'Class'} = $ppro_item->{'Class'}; }
    if (!$new_item->{'SubClass'}) { $new_item->{'SubClass'} = $ppro_item->{'SubClass'}; }
    if (!$new_item->{'VendorCode'}) { $new_item->{'VendorCode'} = $ppro_item->{'VendorCode'}; }
    if ($ppro_item->{'Status'} eq "A") { $new_item->{'Status'} = 'Active'; } else { $new_item->{'Status'} = 'Inactive'; }
    $new_item->{'Cost'} = ($ppro_item->{'StdCost'} != 0?$ppro_item->{'StdCost'}:$ppro_item->{'AvgCost'});
    if (!$new_item->{'CostUOMDescr'}) { $new_item->{'CostUOMDescr'} = $ppro_item->{'PurchUOM'}; }
    
    if ( $product =~ m/^\d\d\d+/ ) {
        $new_item->{'IsRaw'} = 1;
        $new_item->{'Class'} = 'RM';
    }
    else {
        $new_item->{'IsRaw'}    = 0;
        $new_item->{'Class'}    = 'BL';
        $new_item->{'SubClass'} = 'BLEND';
    }

    if ( $validate == 2 ) {
        if ( !$flag ) {
            print "\n\tAdding to Formulator :\t$ppro_item->{'ItemCode'} [$ppro_item->{'Description'}]\n\n";
            my $input = prompt("\tDo you wish to continue [default: Y] ? ");
            if ( ( $input !~ m/^y/i ) && ($input) ) { return 0; }
        }
        $statement = sprintf
"INSERT INTO RawMaterials (ItemCode, Description, Cost, CostUOM, CostUOMDescr, IsRaw, DisplayDensity, VendorCode, ServingSize, ServingUOM, LotTracked, Class, SubClass, Status, AddedBy, Added) VALUES ('%s', '%s', %0.2f, 3, 'KG', %d, 0, '%s', 100, 'mg', 1, '%s', '%s', 'Active', '%s', '%s'); ",$new_item->{'ItemCode'},$new_item->{'Description'},$new_item->{'Cost'},$new_item->{'IsRaw'},$new_item->{'VendorCode'},$new_item->{'Class'},$new_item->{'SubClass'},$username,$day;
        $log->log( " > $statement", 2 );
        my $dbh = FormulaTool->connect_formdb;
        my $sth = $dbh->do($statement) or $log->log("!! Could not do [$statement] \nerror: $!");
        if (($ppro_item->{'AllergenCheck'}>0)&&($form_item->{'AllergenCheck'}>0)) {
            foreach my $count (0 .. $#{ $ppro_item->{'Allergens'} }) {
                $statement2 = sprintf " UPDATE RawAllergens SET ItemCode='%s',AllergenNo=%d,Collateral=%d,[Contains]=%d WHERE ItemCode='%s' AND AllergenNo=%d",$new_item->{'ItemCode'},$count,$ppro_item->{'Allergens'}[$count],$ppro_item->{'Allergens'}[$count],$new_item->{'ItemCode'},$count;
                $log->log( " > $statement2", 2 );
                $sth = $dbh->do($statement2) or $log->log("!! Could not do [$statement2] \nerror: $!");
            }            
        } elsif ($ppro_item->{'AllergenCheck'}>0) {
            foreach my $count (0 .. $#{ $ppro_item->{'Allergens'} }) {
                $statement2 = sprintf " INSERT INTO RawAllergens (ItemCode,AllergenNo,Collateral,[Contains]) VALUES('%s',%d,%d,%d)",$new_item->{'ItemCode'},$count,$ppro_item->{'Allergens'}[$count],$ppro_item->{'Allergens'}[$count];
                $log->log( " > $statement2", 2 );
                $sth = $dbh->do($statement2) or $log->log("!! Could not do [$statement2] \nerror: $!");
            }
        }

        $dbh->disconnect;
    } else {
        if ( !$flag ) {
            print "\n\tProcess Pro :\t$ppro_item->{'ItemCode'} [$ppro_item->{'Description'}]\n\tFormulator :\t$form_item->{'ItemCode'} [$form_item->{'Description'}]\n\n";
            my $input = prompt("\tDo you wish to update Formulator [default: Y] ? ");
            if ( ( $input !~ m/^y/i ) && ($input) ) { return 0; }
        }
        $statement = sprintf "UPDATE RawMaterials SET Description='%s', Cost=%0.2f, CostUOM=3, CostUOMDescr='KG', IsRaw=%d, DisplayDensity=0, VendorCode='%s', LotTracked=1, Class='%s', SubClass='%s', Status='Active', UpdatedBy='%s', Updated='%s' WHERE ItemCode = '%s'",$new_item->{'Description'},$new_item->{'Cost'},$new_item->{'IsRaw'},$new_item->{'VendorCode'},$new_item->{'Class'},$new_item->{'SubClass'},$username,$day,$new_item->{'ItemCode'};
        $log->log( " > $statement", 2 );
        my $dbh = FormulaTool->connect_formdb;
        my $sth = $dbh->do($statement)
          or $log->log("!! Could not do [$statement] \nerror: $!");
        if (($ppro_item->{'AllergenCheck'}>0)&&(defined $form_item->{'Allergens'})) {
            foreach my $count (0 .. $#{ $ppro_item->{'Allergens'} }) {
                $statement2 = sprintf " UPDATE RawAllergens SET ItemCode='%s', AllergenNo=%d, Collateral=%d, [Contains]=%d WHERE ItemCode='%s' AND AllergenNo=%d",$new_item->{'ItemCode'},$count,$ppro_item->{'Allergens'}[$count],$ppro_item->{'Allergens'}[$count],$new_item->{'ItemCode'},$count;
                $log->log( " > $statement2", 2 );
                $sth = $dbh->do($statement2) or $log->log("!! Could not do [$statement2] \nerror: $!");
            }            
        } elsif ($ppro_item->{'AllergenCheck'}>0) {
            foreach my $count (0 .. $#{ $ppro_item->{'Allergens'} }) {
                $statement2 = sprintf " INSERT INTO RawAllergens (ItemCode, AllergenNo, Collateral, [Contains]) VALUES('%s', %d, %d, %d)",$new_item->{'ItemCode'},$count,$ppro_item->{'Allergens'}[$count],$ppro_item->{'Allergens'}[$count];
                $log->log( " > $statement2", 2 );
                $sth = $dbh->do($statement2) or $log->log("!! Could not do [$statement2] \nerror: $!");
            }
        }
        $dbh->disconnect;
    }

    #####

    $log->log(">> ProcessPro Raw Material \t$ppro_item->{'ItemCode'} [$ppro_item->{'Description'}] transferred to Formulator!", 1 );
    return 1;
}

sub xfer_form2f {
#####
    #
    # xfer_form2f - Transfer formula from ProcessPro to Formulator
    #
#####
    local $| = 1;
    my @args    = @_;
    chomp $args[0];
    my $product = $args[0];

    my $statement;
    my $validate = FormulaTool->validate_formula($product);
    if ( $validate == 1 || $validate == 3 || $validate == 5 || $validate == 7 ) { return 0; }
    $log->log( ">>>xfer_form2form $product", 2 );

    my $ppro_form = PProFormula->new($product);
    my $form_form = FormFormula->new($product);
    $ppro_form->{'Description'} =~ s/\s\s+//m;

    #####

    my $complete = xfer_item2f( $product, 1 );
    if ( !$complete ) { return 0; }
    #####

    if (!defined $ppro_form->{'PRWeight'}) {
        $ppro_form->{'PRWeight'} = .1;
    }

    if ( $validate == 2 || $validate == 6) {
        print "\n\tAdding to Formulator :\t$ppro_form->{'FormulaCode'} [$ppro_form->{'Description'}]\n\n";
        my $input = prompt("\tDo you wish to continue [default: Y] ? ");
        if ( ( $input !~ m/^y/i ) && ($input) ) { return 0; }
        $statement = sprintf "INSERT INTO FormulaMaster (FormulaCode, Description, Class, SubClass, Status, AddedBy, Added, DefaultUOM, FormulaTotal, FormulaUOM, CustomerCode) VALUES ('%s', '%s', 'BL', 'BLEND', 'Active', '%s', '%s', 3, %.08f, 'mg', '%s')", $ppro_form->{'FormulaCode'}, $ppro_form->{'Description'}, $username, $day, $ppro_form->{'PRWeight'} * 1000, $ppro_form->{'CustomerCode'};
    } else {
        print "\n\tProcess Pro :\t$ppro_form->{'FormulaCode'} [$ppro_form->{'Description'}]\n\tFormulator :\t$form_form->{'FormulaCode'} [$form_form->{'Description'}]\n\n";
        my $input = prompt("\tDo you wish to update Formulator [default: Y] ? ");
        if ( ( $input !~ m/^y/i ) && ($input) ) { return 0; }
        $statement = sprintf "UPDATE FormulaMaster SET Description='%s', Class='BL', SubClass='BLEND', Status='Active', UpdatedBy='%s', Updated='%s', FormulaTotal=%.08f, FormulaUOM='mg', CustomerCode='%s' WHERE FormulaCode = '%s'", $ppro_form->{'Description'}, $username, $day, $ppro_form->{'PRWeight'} * 1000, $ppro_form->{'CustomerCode'}, $ppro_form->{'FormulaCode'};
    }
    $log->log( " > $statement", 2 );
    my $dbh = FormulaTool->connect_formdb;
    my $sth = $dbh->do($statement)
      or $log->log("!! Could not do [$statement] \nerror: $!");

    #####

    $log->log(" > Dropping Formula Details for $ppro_form->{'FormulaCode'} [$ppro_form->{'Description'}]...", 1 );
    $statement = "DELETE FROM FormulaDetail WHERE FormulaCode like '$product'";
    $log->log(" > $statement", 2);
    $sth       = $dbh->do($statement)
      or $log->log("!! Could not do [$statement] \nerror: $!");

    foreach my $count ( 0 .. $#{ $ppro_form->{'Details'} } ) {
        $log->log(" > Adding $ppro_form->{'Details'}[$count][0] to Formulator details...", 1 );
        $statement = sprintf "INSERT INTO FormulaDetail (FormulaCode, LineType, Code, LineNumber, Quantity, UOM, UOMDescr) VALUES ('%s', 1, '%s', %d, %.08f, 6, 'mg')", $ppro_form->{'FormulaCode'}, $ppro_form->{'Details'}[$count][0], $ppro_form->{'Details'}[$count][1], $ppro_form->{'Details'}[$count][2] * $ppro_form->{'PRWeight'} * 1000;
        $log->log(" > $statement", 2);
        $sth = $dbh->do($statement)
          or $log->log("!! Could not do [$statement] \nerror: $!");
    }

    #####

    $log->log(">> ProcessPro Formula $ppro_form->{'FormulaCode'} [$ppro_form->{'Description'}] transferred to Formulator!", 1 );
    $dbh->disconnect;
    return 1;
}


sub make_mmr {
#####
    #
    # make_mmr - Make MMR for Product ID
    #
#####
    local $| = 1;
    my @args    = @_;
    my $product = $args[0];
    my $brmmr   = $args[1];

    $log->log( ">>>make_mmr $product", 2 );

    my $form_form = FormFormula->new($product);
    my $ppro_form = PProFormula->new($product);
    my $spec_form = SpecFormula->new($product);
    my ( $batchnumber, $index );

    while (1) {
        my (%list);
        print "\nBR and LOT number(s) in ProcessPro (only the past year is displayed):\n";
        print "\n        BR\tLOT NO  \tMFG DATE\n";
        foreach my $count ( 0 .. $#{ $ppro_form->{'Wono'} } ) {
            my @dt       = split( q{ }, $ppro_form->{'Wono'}[$count][3] );
            my @end_date = split( q{-}, $dt[0] );
            my $mfg_date;
            if (   ( $end_date[0] < 2000 )
                && ( $ppro_form->{'Wono'}[$count][4] == 5 ) ) {
                $mfg_date = "- Voided";
            } elsif ( $ppro_form->{'Wono'}[$count][4] == 4 ) {
                $mfg_date = sprintf "%02d/%02d/%04d", $end_date[1], $end_date[2], $end_date[0];
            } else {
                $mfg_date = "- Open";
            }
            $list{ int( $ppro_form->{'Wono'}[$count][1] ) } = $count;
            if (( $end_date[0] > ($date[5] + 1898) )||( $end_date[0] < 2000 )) {
                printf "%s\t%s\t%s\n", $ppro_form->{'Wono'}[$count][1], $ppro_form->{'Wono'}[$count][2], $mfg_date;
            }
        }

        $batchnumber = prompt("\tEnter Batch Record to use <x - cancel>: ");
        
        if ($batchnumber =~ m/^x$/i) {
            $log->log(">> Document canceled!",1);
            return 0;
        } elsif ( (!exists $list{$batchnumber}) || ( int($ppro_form->{'Wono'}[$list{$batchnumber}][2]) !~ m/^\d+$/) ) {
            print "\n\t$batchnumber is invalid!\n";
            next;
        }
        $index = $list{$batchnumber};
        last;
    }

    if ( $ppro_form->{'Wono'}[$index][4] < 4 ) {
        $log->append("WARNING $bell: This BR has not been closed. Changes may still occur!", 1 );
    }

    my $batch = Batch->new($batchnumber);
    $log->log( " > BR used:\t$batch->{'BatchRecord'}",         1 );
    $log->log( " > Total Allocated:\t$batch->{'AllocQty'} Kg", 1 );

    if (!$batch->{'Yield'}) {
        $log->log( "!! Unable to pull data for $batch->{'FormulaCode'} v$batch->{'Version'} BR $batch->{'BatchRecord'} ",1); return 0;
    }

    my $runsize = sprintf "%.3f", $batch->{'RouteSize'} / $batch->{'Yield'} ;
    my $runqty = int($batch->{'AllocQty'} / $runsize);
    my $mod = ( $batch->{'AllocQty'} / $runsize ) - $runqty ;
    my $runrmd = sprintf "%.3f", $mod * $runsize;

    if ( $batch->{'RouteSize'} > 1 ) {
        $log->log( " > Runs:\t" . $runqty . " @ \t$runsize Kg", 1 );
        $log->log( " > Remainder Size:\t$runrmd Kg", 1 );
        while (1) {
            print "\n***\tThis is a multi-run batch.\n\n";
            my $input = prompt("\tDo you want to print the \n\t<W>hole run, <A>ll runs, run #<1-" . ($mod>0?$runqty+1:$runqty) . ">, or <X> - cancel ?" );
            if ( $input =~ m/^x$/i ) {
                $log->log(">> Document canceled!",1);
                last;
            } elsif ( $input =~ m/^a$/i ) {
                foreach my $count (1..$runqty) {
                    write_mmr( $form_form, $ppro_form, $spec_form, $batch, $runsize, $count, $brmmr );
                    next;
                }
                if ( $mod > 0 ) {
                    write_mmr( $form_form, $ppro_form, $spec_form, $batch, $runrmd, $runqty+1, $brmmr );
                    last;
                }
                last;
            } elsif ( $input =~ m/^w$/i ) {
                write_mmr( $form_form, $ppro_form, $spec_form, $batch, $batch->{'AllocQty'}, 0, $brmmr );
                last;
            } elsif ( ($input <= $runqty ) && ($input > 0) ) {
                write_mmr( $form_form, $ppro_form, $spec_form, $batch, $runsize, $input, $brmmr );
                last;
            } elsif ( ($input > $runqty ) && ($runrmd != 0) ) {
                write_mmr( $form_form, $ppro_form, $spec_form, $batch, $runrmd, $runqty+1 , $brmmr );
                last;
            } else {
                print "\n\t$input is invalid!\n";
                next;
            }
        }
    } else {
        write_mmr( $form_form, $ppro_form, $spec_form, $batch, $batch->{'AllocQty'}, 0, $brmmr );
    }
    return 1;
}

sub write_mmr {
#####
    #
    # Write_mmr - Output HTML to file
    #
#####
    local $| = 1;
    my @args      = @_;
    my $form_obj  = $args[0];
    my $ppro_obj  = $args[1];
    my $spec_obj  = $args[2];
    my $batch_obj = $args[3];
    my $runsize   = $args[4];
    my $run       = $args[5];
    my $brmmr     = $args[6];
    my ( @html, $complete, $file );

    if ($brmmr) {
        ( $complete, @html ) = mmr_cover( $form_obj, $ppro_obj, $spec_obj, $batch_obj, $runsize, $run, "MMR" );
    } else {
        ( $complete, @html ) = mmr_cover( $form_obj, $ppro_obj, $spec_obj, $batch_obj, $runsize, $run, "Batch Record" );
    }

    if ( !$complete ) {
        $log->log( "!! Document failed to generate, due to errors ##", 1 );
        return 0;
    }

    my $path = "$ENV{USERPROFILE}\\Desktop\\";
    
    if ($brmmr) {
        $file = sprintf "%d%s %s %s MMR.html", $batch_obj->{'Lot'}, ( $run ? q{.} . $run : q{} ), $ppro_obj->{'FormulaCode'}, $ppro_obj->{'Description'};
    } else {
        $file = sprintf "%d%s %s %s BR.html", $batch_obj->{'Lot'}, ( $run ? q{.} . $run : q{} ), $ppro_obj->{'FormulaCode'}, $ppro_obj->{'Description'};
    }
    $file =~ s|\#||g;
    $file =~ s|/|_|g;
    $file =~ s|^a-zA-Z\d\s[()+,-]|_|g;
    
    my $mmr = LogFile->new( $file, $path );
    $mmr->write( "@html", 2 );
    $log->log( ">> $mmr->{'FilePath'} completed!", 1 );
    return 1;
}


sub mmr_cover {
    local $| = 1;
    my @args      = @_;
    my $form_obj  = $args[0];
    my $ppro_obj  = $args[1];
    my $spec_obj  = $args[2];
    my $batch_obj = $args[3];
    my $runsize   = sprintf "%.3f",$args[4];
    my $run       = $args[5];
    my $type      = $args[6];

    my ( %allergens, @html, $complete, $trurunsize, $truyield );

    my @temp = split( q{ }, $batch_obj->{'ReqDate'} );
    my @due  = split( q{-}, $temp[0] );
    my $duedate = sprintf "%02d/%02d/%04d", $due[1], $due[2], $due[0];

    if (   ( $form_obj->{'ServingDesc'} eq q{} )
        || ( $form_obj->{'ServingSize'} eq q{} ) ) {
        $log->log( '!! There is no Cap/Tab information in Formulator!', 1 );
        return 0;
    }
    my $dose = sprintf "%s %s %s", $form_obj->{'ServingSize'}, $form_obj->{'Appearance'}, $form_obj->{'ServingDesc'};

    my ( $allergen_list, $theo_yield );
    foreach my $type ( keys $ppro_obj->{'AllergenCount'} ) {
        if ( $ppro_obj->{'AllergenCount'}{$type} > 0 ) { $allergen_list .= "$type ," }
    }
    if ( !$allergen_list ) { $allergen_list = 'None'; }

    if ( $ppro_obj->{'PRWeight'} == 0 ) {
        $log->log( " ! Unable to locate $ppro_obj->{'FormulaCode'}.PR!", 1 );
        $theo_yield = 0;
        $truyield = 0;
        $trurunsize = sprintf "%.3f", $runsize;
    } else {
        my $fillweight = $ppro_obj->{'PRWeight'};
        $theo_yield = sprintf "%.0f", $runsize / $fillweight * 1000;
        $theo_yield =~ s/(\d{1,3}?)(?=(\d{3})+$)/$1,/gm;
        if ($ppro_obj->{'Yield'} < 100) {
            $trurunsize = sprintf "%.3f", $runsize * .97 ;
            $truyield = sprintf "%.0f", ($runsize / $ppro_obj->{'PRWeight'} * 1000) * .97;
            $truyield =~ s/(\d{1,3}?)(?=(\d{3})+$)/$1,/gm;
        } else {
            $trurunsize = sprintf "%.3f", $runsize;
            $truyield = sprintf "%.0f", ($runsize / $ppro_obj->{'PRWeight'} * 1000);
            $truyield =~ s/(\d{1,3}?)(?=(\d{3})+$)/$1,/gm;
    }
    }



# # cover_pg1
    $log->log( " > Printing $type Cover Sheet Pg1...", 1 );
    push @html, '
<!DOCTYPE HTML>
<html><head>
<title>' . $type . ' for ' . $batch_obj->{'FormulaCode'} . ' Lot Number ' . $batch_obj->{'Lot'} . ($type eq q{MMR} ? ($run?q{.XX} : q{}) : ($run?q{.}.$run : q{}) ) . '</title>
<style><!--
    @page {
        size: letter landscape;
        margin: 7px 40px 2px 40px;
        padding: 0px; }
    @media print {
        .pagebody {  }
        .noprint { display: none; }
    }
    body {
        counter-reset: page 0;
        font-size:11pt; }
    table	{
        width: 100%;
        height: 8in;
        border: 0px solid black;
        vertical-align: middle; }
    td	{
        height: 15px;
        width: 1vw;
        border: 0px solid black;
        padding-top:1px;
        padding-right:1px;
        padding-left:1px;
        color:windowtext;
        font-size:11pt;
        font-weight:350;
        font-style:normal;
        text-decoration:none;
        font-family:Verdana,Helvetica;
        text-align:general;
        vertical-align:bottom;}
    th	{
        height: 15px;
        width: 1vw;
        border: 0px solid black;
        padding-top:1px;
        padding-right:1px;
        padding-left:1px;
        color:windowtext;
        font-size:12pt;
        font-family:Verdana,Helvetica;
        text-align:general;
        vertical-align:bottom;} 
    li  {
        margin-left: 21px;
        text-indent: -21px;
        font-size:11pt;
        list-style-type: none;}

    .list {
        counter-reset: itemlist;}
    .list li:before {
        content: counter(itemlist, lower-alpha) ") ";
        counter-increment: itemlist;} 
    .heavy	{
        color:black;
        font-size:12pt;
        font-weight:700;
        font-family:Verdana,Helvetica;
        vertical-align: bottom;
        white-space: nowrap; }
    .QA {
        outline: thin solid black; }
    .header {
        top: 0px;
        height: 22px;
        padding-top: 10px;
        font-size: 9pt;
        text-align: center;    }
    .footer {
        bottom: 0px;
        height: 27px;
        padding-top: 0px;
        font-size: 9pt;
        text-align: right; }
    .header:after {
        counter-increment: page;
        content: "' . $type . ' for ' . $batch_obj->{'FormulaCode'} . ($ppro_obj->{'Revision'} > 0? ' rev' . sprintf "%02d", $ppro_obj->{'Revision'} : q{} ) . ($type eq "MMR" ? q{} : ' Lot Number ' . $batch_obj->{'Lot'} . ( $run ? "\.XX" : q{} ) ). ' (' . $runsize . ' Kg)"; }
    .footer:after {
        content: "v' .$VERSION. ' printed ' . $day . ' " " page " counter(page); }
--></style>';
    push @html, '</head><body>';
    
    push @html, $main_tableh;
    if ($type eq "MMR") {
        push @html, $single_blank;
        push @html, '
 <tr style="height: 30px">
  <td colspan="100" align="center"><h2>' . $type . ' for ' . $batch_obj->{'FormulaCode'} . ' Lot Number ' . $batch_obj->{'Lot'} . ($run?q{.XX} : q{}) . (sprintf " - %d Kg", $runsize) . ' <img src="http://icons.iconarchive.com/icons/icojam/onebit-4/32/printer-icon.png" width="30" onclick="window.print();" style="cursor:pointer;" class="noprint"></h2></td>
 </tr>';
        push @html, $single_blank;
    } else {
        push @html, '<tr class="single"><td colspan="80">&nbsp;</td><td colspan="20" align="right">'. ($run<2?'SO#&nbsp;_______________' : q{}) .'</td></tr>';
        push @html, '
 <tr style="height: 30px">
  <td colspan="100" align="center"><h2>' . $type . ' for ' . $batch_obj->{'FormulaCode'} . ' Lot Number ' . $batch_obj->{'Lot'} . ($run ?  "\." . $run : q{}) . ' <img src="http://icons.iconarchive.com/icons/icojam/onebit-4/32/printer-icon.png" width="30" onclick="window.print();" style="cursor:pointer;" class="noprint"></h2></td>
 </tr>';
        push @html, '
 <tr>
  <td colspan="80">&nbsp;</td> 
  <td colspan="20" align="right">'. ($run>1?"&nbsp;":"Tier&nbsp;2#&nbsp;_______________") .'</td>
 </tr>';
    }
#6 lines
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="15" align="left">Product Master</td>
  <td colspan="40" align="left">: ' . $batch_obj->{'FormulaCode'} . '</td>
  <td colspan="25">&nbsp;</td>  
  <td colspan="20" align="right">'. (($type eq "MMR")||($run>1)?"&nbsp;":"Tier&nbsp;3#&nbsp;_______________") .'</td>
 </tr>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="15" align="left">Product Name</td>
  <td colspan="40" align="left">: ' . $ppro_obj->{'Description'} . '</td>
  <td colspan="12">&nbsp;</td>
  '. (($type eq "MMR")||($run>1)? '<td colspan="20">&nbsp;</td>' : '<td colspan="8" align="right">Prior Lot</td>
  <td colspan="12" valign="top">: '
      . ( $batch_obj->{'PriorLot'} ? $batch_obj->{'PriorLot'} : "n/a" ) . '</td>') . '    
 </tr>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="15" align="left">Customer Name</td>
  <td colspan="40" align="left">: ' . $form_obj->{'CustomerName'} . '</td>
  <td colspan="12" align="right">&nbsp;</td>
  '. ($type eq "MMR"?'<td colspan="20">&nbsp;</td>': '<td colspan="8" align="right">Due Date</td>'. '
  <td colspan="12" align="left">'.($type eq "MMR"?'&nbsp;':': ' . $duedate)) . '</td>
 </tr>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="15" align="left">Serving Description</td>
  <td colspan="40" align="left">: ' . $dose . '</td>
  <td colspan="12" align="right">&nbsp;</td>
  <td colspan="8" align="right" style="vertical-align:top">Batch Weight</td>
  <td colspan="12" align="left" style="vertical-align:top">: ' . $runsize . ' Kg'. ($type eq "MMR" ?'<br>:&nbsp;('. $trurunsize . ' Kg)':q{}). '</td>
 </tr>';
    push @html, $single_blank;
    push @html, ' 
 <tr>
  <td colspan="15" align="left">Certified Organic?</td>
  <td colspan="5" align="center">Yes</td>
  <td colspan="5" align="center">No</td>
  <td colspan="42" align="right"></td>';
      if ($theo_yield) {
        my $uom;
        if ( $form_obj->{'ServingDesc'} =~ m/Tab/ ) {
            $uom = " Tablets";
        } elsif ( $form_obj->{'ServingDesc'} =~ m/Cap/ ) {
            $uom = " Capsules";
        } else {
            $uom = " Kg"
        }
        push @html, '<td colspan="8" align="right" style="vertical-align:top">Batch Qty</td>
        <td colspan="12" align="left" style="vertical-align:top">: ' . $theo_yield . $uom . ($type eq "MMR" ?'<br>:&nbsp;('. $truyield.$uom. ')':q{}) . '</td>';
      } else {
        push @html, '<td colspan="20" align="right">&nbsp;</td>';
      }
    push @html, '
 </tr>';
 #10 lines
    push @html, $single_blank;
    push @html, ' 
 <tr>
  <td colspan="15" align="left">Bulk Order?</td>
  <td colspan="5" align="center">Yes</td>
  <td colspan="5" align="center">No</td>
  <td colspan="75" align="right">&nbsp;</td>
 </tr>';
    push @html, $single_blank;
#5 lines
    push @html, '
  <tr style="height:30px">
   <td colspan="15" align="left" class="heavy" style="vertical-align:top; font-size: 18pt">Allergens</td>
   <td colspan="85" align="left" class="heavy" style="vertical-align:top; font-size: 18pt;">: ';
    if (scalar(keys %{ $ppro_obj->{'Allergens'} }) < 1) {
        push @html, '&nbsp;None<br></td></tr>';
    } else {
        foreach my $raw (keys %{ $ppro_obj->{'Allergens'} }) {
            push @html, $raw .' '. $ppro_obj->{'Allergens'}{$raw} .', ';
        }
    }
    push @html, '</td>
 </tr>';
        push @html, $single_blank;
        push @html, $single_blank;
        push @html, $single_blank;
        push @html, $single_blank;
        push @html, $single_blank;
        push @html, $single_blank;
        push @html, $single_blank;
        push @html, $single_blank;
        push @html, $single_blank;
        push @html, $single_blank;
        push @html, $single_blank;
#12 lines    
    if (($run>1) or ($type eq "MMR"))  {
        push @html, $single_blank;
        push @html, $single_blank;
        push @html, $single_blank;
        push @html, $single_blank;
        push @html, $single_blank;
        push @html, $single_blank;
        push @html, $single_blank;
    } else {
        push @html, '        
<tr>
    <td colspan="100" align="right"><span style="font-size:75%">(eval: '.$ppro_obj->{'LastCost'}.')</span></td>
</tr>';
        push @html, '
<tr>
    <td colspan="30" class="heavy">Samples Collected</td>
    <td colspan="70" class="heavy">&nbsp;</td>
</tr>';
        push @html, '
<tr>
<td colspan="30" class=q{}>&nbsp;Potency (Bulk Final Dosage form)</td>
<td colspan="22">&nbsp;</td>
<td colspan="13" align="center" class="heavy">Date:&nbsp;_________</td>
<td colspan="13" align="center" class="heavy">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
<td colspan="13" align="center" class="heavy">Initials:&nbsp;_________</td>
<td colspan="9" align="right">QA-427</td>
</tr>';
        push @html, $single_blank;
        push @html, '
<tr>
<td colspan="30">&nbsp;Heavy Metals, Microbial and Retained<br>&nbsp;(Finished Package form only)</td>
<td colspan="22">&nbsp;</td>
<td colspan="13" align="center" class="heavy">Date:&nbsp;_________</td>
<td colspan="13" align="center" class="heavy">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
<td colspan="13" align="center" class="heavy">Initials:&nbsp;_________</td>
<td colspan="9" align="right">QA-427</td>
</tr>';    
    }    
#4 lines
    if ($type eq "MMR") {
        push @html, $single_blank;
        push @html, '
 <tr>
  <td colspan="30" class="heavy">MMR Issued/Verified by</td>
  <td colspan="22">&nbsp;</td>
  <td colspan="13" align="center" class="heavy">Date:&nbsp;_________</td>
  <td colspan="13" align="center" class="heavy">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="13" align="center" class="heavy">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">&nbsp;</td>
 </tr>';
        push @html, $single_blank;
        push @html, $single_blank;
        push @html, $single_blank;
    } else { 
        push @html, $single_blank;
        push @html, '<tbody class="QA">' . $single_blank;
        push @html, '
 <tr>
  <td colspan="30" class="heavy">QC Release to Production</td>
  <td colspan="22">&nbsp;</td>
  <td colspan="13" align="center" class="heavy">Date:&nbsp;_________</td>
  <td colspan="13" align="center" class="heavy">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="13" align="center" class="heavy">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">QA-445</td>
 </tr>';
        push @html, $single_blank;
        push @html, '
 <tr>
  <td colspan="30" class="heavy">QC Release for Packaging</td>
  <td colspan="22">&nbsp;</td>
  <td colspan="13" align="center" class="heavy">Date:&nbsp;_________</td>
  <td colspan="13" align="center" class="heavy">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="13" align="center" class="heavy">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">QA-445</td>
 </tr>';
        push @html, $single_blank;
        push @html, '
 <tr>
  <td colspan="30" class="heavy">QC Release for Shipment</td>
  <td colspan="22">&nbsp;</td>
  <td colspan="13" align="center" class="heavy">Date:&nbsp;_________</td>
  <td colspan="13" align="center" class="heavy">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="13" align="center" class="heavy">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">QA-445</td>
 </tr>';     
        push @html, $single_blank.'</tbody>';
        push @html, $single_blank; 
    } 
    push @html, $single_blank;
    push @html, $main_tablef;
#12 lines
    ( $complete, @html ) =
      mmr_pg1( $form_obj, $ppro_obj, $spec_obj, $batch_obj, $runsize, $run, @html );
    return ( $complete, @html );
}

sub mmr_pg1 {
#####
    #
    # Section 1 of MMR - Weighing and Blending
    #
#####
    my @args      = @_;
    my $form_obj  = shift @args;
    my $ppro_obj  = shift @args;
    my $spec_obj  = shift @args;
    my $batch_obj = shift @args;
    my $runsize   = sprintf "%.3f",shift @args;
    my $run       = shift @args;
    my @html      = @args;
    my $preblend  = 0;
    my $pbcount   = 0;
    my $rawcount  = $#{ $ppro_obj->{'Details'} } +1;
    my $pbtheo    = $runsize * .05;    
    foreach my $count ( 0 .. $#{ $ppro_obj->{'Details'} } ) {
        my $portion = $ppro_obj->{'Details'}[$count][2] * $runsize;
        if ( $portion <= 1 ) { 
            $preblend += $portion; $pbcount++; }
    }
    $log->log( " > Preblend items: $pbcount", 1 );    
    $log->log( " > Preblend portion: $preblend Kg", 1 );

# # w&b_pg2
    $log->log( " > Printing Weigh Room Clean Check Pg2...", 1 );
    push @html, $main_tableh;
    push @html, '
 <tr>
  <td colspan="100" align="left" class="heavy">Weigh Room Clean Check</td>
 </tr>';
    push @html, $single_blank; 
    push @html, '
 <tr>
  <td colspan="30">Screen Cleaned</td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="23">&nbsp;</td>
  <td colspan="11" align="center">Date:&nbsp;_________</td>
  <td colspan="11" align="center">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">MF-274</td>
 </tr>';
    push @html, $single_blank; 
    push @html, '
 <tr>
  <td colspan="30">Scoops Cleaned</td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="23">&nbsp;</td>
  <td colspan="11" align="center">Date:&nbsp;_________</td>
  <td colspan="11" align="center">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">MF-274</td>
 </tr>';
    push @html, $single_blank; 
    push @html, '
 <tr>
  <td colspan="30">Scale #________ Cleaned</td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="23">&nbsp;</td>
  <td colspan="11" align="center">Date:&nbsp;_________</td>
  <td colspan="11" align="center">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">MF-274</td>
 </tr>';
    push @html, $single_blank; 
    push @html, '
 <tr>
  <td colspan="30">Room #________ Cleaned</td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="23">&nbsp;</td>
  <td colspan="11" align="center">Date:&nbsp;_________</td>
  <td colspan="11" align="center">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">MF-215</td>
 </tr>';
    push @html, $single_blank; 
    push @html, '
 <tr>
  <td colspan="30">Organic Rinse?</td>
  <td colspan="5" align="center">Yes</td>
  <td colspan="5" align="center">N/A</td>
  <td colspan="51" align="right">&nbsp;</td>
  <td colspan="9" align="right">MF-274</td>
 </tr>';
    push @html, $single_blank; 
    push @html, '
 <tr>
  <td colspan="30">BR Notes Read?</td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="23">&nbsp;</td>
  <td colspan="11" align="center">Date:&nbsp;_________</td>
  <td colspan="11" align="center">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">&nbsp;</td>
 </tr>';
#14 lines
    push @html, $single_blank; 
    push @html, '<tbody class="QA">' . $single_blank;
    push @html, '
 <tr>
  <td colspan="30" class="heavy">QC Approval Only</td>
  <td colspan="28">&nbsp;</td>
  <td colspan="11" align="center" class="heavy">Date:&nbsp;_________</td>
  <td colspan="11" align="center" class="heavy">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center" class="heavy">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">&nbsp;</td>
 </tr>';
    push @html, $single_blank . '</tbody>'; 
    push @html, $single_blank; 
    push @html, $single_blank; 
    push @html, $single_blank; 
    push @html, $single_blank; 
    push @html, $single_blank; 
    push @html, $single_blank; 
    push @html, $single_blank; 
    push @html, $single_blank;
#13 lines
    push @html, $single_blank; 
    push @html, '
 <tr>
  <td colspan="100" align="left" class="heavy">Portable Equipment Used</td>
 </tr>';
    push @html, $single_blank; 
    push @html, '
 <tr>
  <td colspan="30">Equipment&nbsp;Used (circle&nbsp;one):</td>
  <td colspan="15" align="center" class="heavy">Colton Mill</td>
  <td colspan="15" align="center" class="heavy">Hammer Mill</td>
  <td colspan="15" align="center" class="heavy">Chilsonator</td>
  <td colspan="25">&nbsp;</td>
 </tr>';
    push @html, $single_blank; 
    push @html, '
 <tr>
  <td colspan="30">Room #________ Cleaned</td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="23">&nbsp;</td>
  <td colspan="11" align="center">Date:&nbsp;_________</td>
  <td colspan="11" align="center">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">MF-215</td>
 </tr>';
    push @html, $single_blank; 
    push @html, '
 <tr>
  <td colspan="30">Equipment Cleaned</td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="23">&nbsp;</td>
  <td colspan="11" align="center">Date:&nbsp;_________</td>
  <td colspan="11" align="center">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">MF-274</td>
 </tr>';  
    push @html, $single_blank; 
    push @html, '
 <tr>
  <td colspan="30">Organic Rinse?</td>
  <td colspan="5" align="center">Yes</td>
  <td colspan="5" align="center">N/A</td>
  <td colspan="51" align="right">&nbsp;</td>
  <td colspan="9" align="right">MF-274</td>
</tr>';
    push @html, $single_blank; 
    push @html, '
 <tr>
  <td colspan="30">Incoming Weight</td>
  <td colspan="5" align="center">_________&nbsp;Kg</td>
  <td colspan="23">&nbsp;</td>
  <td colspan="11" align="center">Date:&nbsp;_________</td>
  <td colspan="11" align="center">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center">Initials:&nbsp;_________</td>
  <td colspan="9" align="right"></td>
 </tr>';  
    push @html, $single_blank; 
    push @html, '
 <tr>
  <td colspan="30">Outgoing Weight</td>
  <td colspan="5" align="center">_________&nbsp;Kg</td>
  <td colspan="23">&nbsp;</td>
  <td colspan="11" align="center">Date:&nbsp;_________</td>
  <td colspan="11" align="center">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center">Initials:&nbsp;_________</td>
  <td colspan="9" align="right"></td>
 </tr>';  
 #14 lines
    push @html, $single_blank; 
    push @html, '<tbody class="QA">' . $single_blank;
    push @html, '
 <tr>
  <td colspan="30" class="heavy">QC Approval Only</td>
  <td colspan="28">&nbsp;</td>
  <td colspan="11" align="center" class="heavy">Date:&nbsp;_________</td>
  <td colspan="11" align="center" class="heavy">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center" class="heavy">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">&nbsp;</td>
 </tr>';
    push @html, $single_blank . '</tbody>'; 
    push @html, $single_blank;
    push @html, $single_blank; 
    push @html, $single_blank; 
    push @html, $single_blank; 
    push @html, $single_blank; 
    push @html, $single_blank;
    push @html, $main_tablef;
#15 lines

    $log->log( " > Pre-blend Verifications:\n\trawcount $rawcount\n\tpbcount $pbcount\n\tpbtheo $pbtheo\n\tpreblend $preblend", 2 );
    if (($rawcount-$pbcount>0)&&($pbcount>0)) {
        if ($pbtheo < $preblend) { $pbtheo = $preblend; }
        &write_preblend( $ppro_obj, $batch_obj, $runsize, $run, 1  );
    }
# # pb_pg3
    $log->log( " > Printing Pre-blend Verification Pg3...", 1 );
    push @html, $main_tableh;
    push @html, '
 <tr>
  <td colspan="91" class="heavy">Pre-Blend Verification</td>
  <td colspan="9" align="right">MF-214</td>
 </tr>';   
    push @html, '
 <tr style="height: 45px">
    <td colspan="100">
        <li style="list-style: disc inside">All ingredients weighing 1 Kg or less, are to be added to the Pre-Blend.</li>
        <li style="list-style: disc inside">Ingredients weighing 100 gm or less, must be wieghed using the Gram scale.</li>
        <li style="list-style: disc inside">Total Pre-Blend weight must weigh at least 5% of total blend weight</li></td>
 </tr>';
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="100" class="heavy">Total Pre-Blend Weight (Theoretical: ' . ((($rawcount-$pbcount>0)&&($pbcount>0))?(sprintf "%.3f&nbsp;Kg", $pbtheo):'No Pre-Blend') . ')</td>
 </tr>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="30">Scale #________ Used</td>
  <td colspan="28" align="center">&nbsp;</td>
  <td colspan="11" align="center">Date:&nbsp;________</td>
  <td colspan="11" align="center">Time:&nbsp;________<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">&nbsp;</td>  
 </tr>';
    push @html, $single_blank;
    push @html, $single_blank; 
    push @html, '
 <tr>
  <td colspan="100" class="heavy">Weight Verifications (Theoretical: '
      . $runsize . '&nbsp;Kg)</td>
 </tr>';
    push @html, '
  <tr style="height: 45px">
   <td colspan="25">&nbsp;</td>
   <th colspan="12">Tare</th>
   <th colspan="12">Net Weight</th>
   <th colspan="12">Date</th>
   <th colspan="12">Time</th>
   <th colspan="18">Initials<br>(weighed/verified)</th>
   <td colspan="9">&nbsp;</td>
  </tr>';
    push @html, '
 <tr style="height: 45px">
  <td colspan="10">&nbsp;</td>
  <td colspan="15" align="center">Pre-Blend</td>
  <td colspan="12" align="center">______________&nbsp;Kg</td>
  <td colspan="12" align="center">______________&nbsp;Kg</td>
  <td colspan="12" align="center">______________</td>
  <td colspan="12" align="center">_______:_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="18" align="center">__________/__________</td>
  <td colspan="9">&nbsp;</td>
 </tr>';
    foreach my $count ( 1 .. 12 ) {
        push @html, '
 <tr style="height: 45px">
  <td colspan="10">&nbsp;</td>
  <td colspan="15" align="center">Drum ' . $count . '</td>
  <td colspan="12" align="center">______________&nbsp;Kg</td>
  <td colspan="12" align="center">______________&nbsp;Kg</td>
  <td colspan="12" align="center">______________</td>
  <td colspan="12" align="center">_______:_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="18" align="center">__________/__________</td>
  <td colspan="9">&nbsp;</td>
 </tr>';
    }
    push @html, $single_blank;
    push @html, '
  <tr>
   <td class="heavy" colspan="25">Total Weight</td>
  <td colspan="12" align="center" class="heavy">______________&nbsp;Kg</td>
   <td colspan="63">&nbsp;</td>
  </tr>';
    push @html, $single_blank;
    push @html, '
  <tr>
   <td colspan="30" class="heavy">Total Weight<br>must be between</td>
   <td colspan="10" class="heavy" align="center" style="vertical-align: bottom">'
      . ( sprintf "%.3f", $runsize * .995 <= 0 ? $runsize : $runsize * .995 )
      . '&nbsp;Kg</td>
   <td colspan="5" align="center">and</td>
   <td colspan="10" class="heavy" align="center" style="vertical-align: bottom">'
      . ( sprintf "%.3f", $runsize * 1.005 ) . '&nbsp;Kg</td>
   <td colspan="14" align="right">(&plusmn;0.5%)&nbsp;Conforms?</td>
   <td colspan="11" align="center">Yes</td>
   <td colspan="11" align="center">No</td>
   <td colspan="9">&nbsp;</td>
 </tr>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="58">If Total Weight is outside of the range, contact Quality Control to address deviation.</td>
  <td colspan="11" class="heavy">QC Deviation:</td>
  <td colspan="11" align="center"># _________</td>
  <td colspan="11" align="center">N/A</td>
  <td colspan="9" align="right">QA-487</td>
 </tr>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="30" class="heavy">QC Approval Only</td>
  <td colspan="28">&nbsp;</td>
  <td colspan="11" align="center" class="heavy">Date:&nbsp;_________</td>
  <td colspan="11" align="center" class="heavy">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center" class="heavy">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">&nbsp;</td>
 </tr>';
    push @html, $main_tablef;
  
  
# # w&b_pg4
    $log->log( " > Printing Blend Area Clean Check Pg4...", 1 );
    push @html, $main_tableh;
    push @html, '
 <tr>
  <td colspan="100" align="left" class="heavy">Blender Room Clean Check</td>
 </tr>';
    push @html, $single_blank; 
    push @html, '
 <tr>
  <td colspan="30">Blender Cleaned</td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="23">&nbsp;</td>
  <td colspan="11" align="center">Date:&nbsp;_________</td>
  <td colspan="11" align="center">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">MF-217</td>  
 </tr>';
    push @html, $single_blank; 
    push @html, '
 <tr>
  <td colspan="30">Scale #________ Cleaned</td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="23">&nbsp;</td>
  <td colspan="11" align="center">Date:&nbsp;_________</td>
  <td colspan="11" align="center">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">MF-274</td>
 </tr>';
    push @html, $single_blank; 
    push @html, '
 <tr>
  <td colspan="30">Room #________ Cleaned</td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="23">&nbsp;</td>
  <td colspan="11" align="center">Date:&nbsp;_________</td>
  <td colspan="11" align="center">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">MF-217</td>
 </tr>';
    push @html, $single_blank; 
    push @html, '
 <tr>
  <td colspan="30">Organic Rinse?</td>
  <td colspan="5" align="center">Yes</td>
  <td colspan="5" align="center">N/A</td>
  <td colspan="51" align="right">&nbsp;</td>
  <td colspan="9" align="right">MF-274</td>
 </tr>';
    push @html, $single_blank; 
    push @html, '<tbody class="QA">' . $single_blank;
    push @html, '
 <tr>
  <td colspan="30" class="heavy">QC Approval Only</td>
  <td colspan="28">&nbsp;</td>
  <td colspan="11" align="center" class="heavy">Date:&nbsp;_________</td>
  <td colspan="11" align="center" class="heavy">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center" class="heavy">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">&nbsp;</td>
 </tr>';
    push @html, $single_blank . '</tbody>';  
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank; 
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank; 
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank; 
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank; 
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank; 
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;  
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;  
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $main_tablef;


# # w&b_pg5
    $log->log( " > Printing Blending Signoffs Pg5...", 1 );
    push @html, $main_tableh;
    if (($rawcount-$pbcount>0)&&($pbcount>0)) {    # preblend needed if > 2 items in formula, and at least 1 preblend material
        push @html, '
 <tr>
  <td colspan="100" class="heavy">Pre-Blend Blending (Theoretical: '. sprintf("%0.3f",$pbtheo) .' Kg)</td>
 </tr>';
        push @html, $single_blank;
        push @html, '
 <tr>
  <td colspan="30" class="heavy">Blender #__________</td>
  <td colspan="15" align="center" class="heavy">Blend&nbsp;Time:&nbsp;_______&nbsp;Minutes</td>
  <td colspan="13">&nbsp;</td>
  <td colspan="11" align="center" class="heavy">Date:&nbsp;_________</td>
  <td colspan="11" align="center" class="heavy">Time:&nbsp;________<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center" class="heavy">Quality:&nbsp;_________</td>
   <td colspan="9">&nbsp;</td>
 </tr>';
        push @html, '
 <tr style="height:30px">
  <td colspan="30">Blended _______ Minutes</td>
  <td colspan="5" align="center">' . $check_box . '</td>
  <td colspan="23">&nbsp;</td>
  <td colspan="11" align="center">Date:&nbsp;_________</td>
  <td colspan="11" align="center">Time:&nbsp;________<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">&nbsp;</td>
 </tr>';
        push @html, '
 <tr style="height:30px">
  <td colspan="30">Blender Cleaned</td>
  <td colspan="5" align="center">' . $check_box . '</td>
  <td colspan="23">&nbsp;</td>
  <td colspan="11" align="center">Date:&nbsp;_________</td>
  <td colspan="11" align="center">Time:&nbsp;________<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">MF-217</td>
 </tr>';
    } else {
       $log->log( " > No Pre-blend!", 1 );
        push @html, '
 <tr>
  <td colspan="100" class="heavy">Pre-Blend Blending (Theoretical: No Pre-Blend)</td>
 </tr>';
        push @html, $single_blank;
        push @html, $single_blank;
        push @html, $single_blank;        
        push @html, $single_blank;        
    }
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="100" class="heavy">Main Blend Blending - Main Blend PLUS Pre-Blend (Theoretical: '
      . $runsize . ' kg)</td>
 </tr>';
        push @html, $single_blank;
        push @html, '
 <tr>
  <td colspan="30" class="heavy">Blender #__________</td>
  <td colspan="15" align="center" class="heavy">Blend&nbsp;Time:&nbsp;_______&nbsp;Minutes</td>
  <td colspan="13">&nbsp;</td>
  <td colspan="11" align="center" class="heavy">Date:&nbsp;_________</td>
  <td colspan="11" align="center" class="heavy">Time:&nbsp;________<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center" class="heavy">Quality:&nbsp;_________</td>
   <td colspan="9">&nbsp;</td>
 </tr>';
        push @html, '
 <tr style="height:30px">
  <td colspan="30">Blended _______ Minutes</td>
  <td colspan="5" align="center">' . $check_box . '</td>
  <td colspan="23">&nbsp;</td>
  <td colspan="11" align="center">Date:&nbsp;_________</td>
  <td colspan="11" align="center">Time:&nbsp;________<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">&nbsp;</td>
 </tr>';
        push @html, '
 <tr style="height:30px">
  <td colspan="30">Blender Cleaned</td>
  <td colspan="5" align="center">' . $check_box . '</td>
  <td colspan="23">&nbsp;</td>
  <td colspan="11" align="center">Date:&nbsp;_________</td>
  <td colspan="11" align="center">Time:&nbsp;________<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">MF-217</td>
 </tr>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="100" class="heavy">Post Blend Weight Verification</td>
</tr>';
    push @html, '
 <tr>
  <td colspan="91">After the Blending is complete, remove powder from the blender and verify the Total Blend Weight</td>
  <td colspan="9" align="right">&nbsp;</td>
 </tr>';
    push @html, '
 <tr>
  <td colspan="100" class="heavy">Total Weight Blended (Theoretical: '
      . $runsize . ' kg)</td>
 </tr>';
    push @html, '
  <tr style="height: 37px">
   <td colspan="25">&nbsp;</td>
   <th colspan="12">Tare</th>
   <th colspan="12">Net Weight</th>
   <th colspan="12">Date</th>
   <th colspan="12">Time</th>
   <th colspan="12">Initials</th>
   <td colspan="15">&nbsp;</td>
  </tr>';
    foreach my $count ( 1 .. 12 ) {
        push @html, '
 <tr style="height: 37px">
  <td colspan="10">&nbsp;</td>
  <td colspan="15" align="center">Drum ' . $count . '</td>
  <td colspan="12" align="center">______________&nbsp;Kg</td>
  <td colspan="12" align="center">______________&nbsp;Kg</td>
  <td colspan="12" align="center">______________</td>
  <td colspan="12" align="center">_______:_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="12" align="center">______________</td>
  <td colspan="15">&nbsp;</td>
 </tr>';
    }
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="25" class="heavy">Total Weight Blended</td>
  <td class="heavy" colspan="12" align="center">______________&nbsp;Kg</td>
  <td class="heavy" colspan="63">&nbsp;</td>
 </tr>';
    push @html, $single_blank;
    my $lo_range = sprintf "%.3f", ( $runsize * .99 );
    my $hi_range = sprintf "%.3f", ( $runsize * 1.01 );
    push @html, '
 <tr>
  <td colspan="30" class="heavy">Total Weight Blended<br>must be between</td>
  <td colspan="10" class="heavy" style="vertical-align: bottom;" align="center">' . $lo_range . '&nbsp;Kg</td>
  <td colspan="5" align="center">and</td>
  <td colspan="10" class="heavy" style="vertical-align: bottom;" align="center">' . $hi_range . '&nbsp;Kg</td>
  <td colspan="14" align="right">(&plusmn;1&#37;)&nbsp;Conforms?</td>
  <td colspan="11" align="center">Yes</td>
  <td colspan="11" align="center">No</td>
  <td colspan="9">&nbsp;</td>
 </tr>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="58">If Total Weight is outside of the range, contact Quality Control to address deviation.</td>
  <td colspan="11" class="heavy">QC Deviation:</td>
  <td colspan="11" align="center"># _________</td>
  <td colspan="11" align="center">N/A</td>
  <td colspan="9" align="right">QA-487</td>
 </tr>';
 
    push @html, $single_blank;
    my $verb;
    if ( $form_obj->{'ServingDesc'} =~ m/Tab/ ) {
        $verb = "reconciled and is released to be tableted";  # Tableting Signoffs
    } elsif ( $form_obj->{'ServingDesc'} =~ m/Cap/ ) {
        $verb = "reconciled and is released to be encapsulated";  # Encapsulation Signoffs
    } else {
        $verb = "reconciled and is released to bottling";   # Powder Signoffs 
    }
    push @html, '
 <tr>
  <td colspan="100" class="heavy">This lot has been properly weighed, blended, ' . $verb . ':</td>
 </tr>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="30" class="heavy">QC Approval Only</td>
  <td colspan="28">&nbsp;</td>
  <td colspan="11" align="center" class="heavy">Date:&nbsp;_________</td>
  <td colspan="11" align="center" class="heavy">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center" class="heavy">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">&nbsp;</td>
 </tr>';
    push @html, $main_tablef;

    my $complete;
    if ( $form_obj->{'ServingDesc'} =~ m/Tab/ ) {
        ( $complete, @html ) = mmr_tab( $form_obj, $ppro_obj, $spec_obj, $batch_obj, $runsize, @html );        # Tableting Signoffs
    } elsif ( $form_obj->{'ServingDesc'} =~ m/Cap/ ) {
        ( $complete, @html ) = mmr_cap( $form_obj, $ppro_obj, $spec_obj, $batch_obj, $runsize, @html );        # Encapsulation Signoffs
    } else {
        ( $complete, @html ) = mmr_pow( $form_obj, $ppro_obj, $spec_obj, $batch_obj, $runsize, @html );        # Powder Signoffs
    }
    #( $complete, @html ) = mmr_pkg( $form_obj, $ppro_obj, $spec_obj, $batch_obj, $runsize, @html );            # Packaging Signoffs
    push @html, '
</body>';
    if ( !$complete ) { return 0; }
    return ( 1, @html );
}


sub mmr_cap {
#####
    #
    # Secrtion 2 of MMR - Encapsulation
    #
#####
    my @args      = @_;
    my $form_obj  = shift @args;
    my $ppro_obj  = shift @args;
    my $spec_obj  = shift @args;
    my $batch_obj = shift @args;
    my $runsize   = shift @args;
    my @html      = @args;
    my ( $cap_add, $cap_adj );

    if ( $form_obj->{'ServingSize'} =~ m/^Size 3\w*/ ) { $cap_add = 50; $cap_adj = 3; }
    if ( $form_obj->{'ServingSize'} =~ m/^Size 2\w*/ ) { $cap_add = 60; $cap_adj = 4; }
    if ( $form_obj->{'ServingSize'} =~ m/^Size 1\w*/ ) { $cap_add = 75; $cap_adj = 5; }
    if ( $form_obj->{'ServingSize'} =~ m/^Size 0\w*/ ) { $cap_add = 100; $cap_adj = 6; }
    if ( $form_obj->{'ServingSize'} =~ m/^Size 0el\w*/ ) { $cap_add = 110; $cap_adj = 7; }
    if ( $form_obj->{'ServingSize'} =~ m/^Size 00\w*/ ) {$cap_add = 120; $cap_adj = 7; }
    if ( $form_obj->{'ServingSize'} =~ m/^Size 00el\w*/ ) { $cap_add = 130; $cap_adj = 10; }
    if ( $form_obj->{'ServingSize'} =~ m/^Size 000\w*/ ) { $cap_add = 165; $cap_adj = 10; }
    if ( $form_obj->{'ServingSize'} !~ m/^Size/ ) {
        $log->log( "!! Capsule Size is not set up correctly in Formulator!", 1 );
        return 0;
    }

    if ( !$ppro_obj->{'PRWeight'} ) {
        $log->log( "!! Unable to locate Fill Weight in ProcessPro!", 1 );
        return 0;
    }

    my $lo_range   = sprintf "%.3f", $runsize * .99;
    my $hi_range   = sprintf "%.3f", $runsize * 1.01;
    my $fillweight = sprintf "%.0f", $ppro_obj->{'PRWeight'} * 1000;
    my $trg_weight = sprintf "%.0f", $fillweight + $cap_add;
    my $lo_weight  = sprintf "%.0f", ( $fillweight + $cap_add ) * 0.95;
    my $hi_weight  = sprintf "%.0f", ( $fillweight + $cap_add ) * 1.05; # + $cap_adj;
    my $theo_yield = sprintf "%.0f", $runsize / $fillweight * 1000000;
    my $lo_yield   = sprintf "%.0f", $theo_yield * .970;
    my $hi_yield   = sprintf "%.0f", $theo_yield * 1.030;
    $trg_weight =~ s/(\d{1,3}?)(?=(\d{3})+$)/$1,/gm;
    $lo_weight =~ s/(\d{1,3}?)(?=(\d{3})+$)/$1,/gm;
    $hi_weight =~ s/(\d{1,3}?)(?=(\d{3})+$)/$1,/gm;
    $theo_yield =~ s/(\d{1,3}?)(?=(\d{3})+$)/$1,/gm;
    $lo_yield =~ s/(\d{1,3}?)(?=(\d{3})+$)/$1,/gm;
    $hi_yield =~ s/(\d{1,3}?)(?=(\d{3})+$)/$1,/gm;

 # # encap_pg6
    $log->log( " > Printing Encapsulation Clean Check Pg6...", 1 );
    push @html, $main_tableh;
    push @html, '
 <tr>
  <td colspan="100" align="left" class="heavy">Encapsulation Area Clean Check</td>
 </tr>';
    push @html, $single_blank;
    push @html, ' 
 <tr>
  <td colspan="30">Equipment #_______ Cleaned</td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="23">&nbsp;</td>
  <td colspan="11" align="center">Date:&nbsp;_________</td>
  <td colspan="11" align="center">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">MF-275'."\n".'MF-276</td>
 </tr>';
    push @html, $single_blank; 
    push @html, '
 <tr>
  <td colspan="30">Room #_______ Cleaned</td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="23">&nbsp;</td>
  <td colspan="11" align="center">Date:&nbsp;_________</td>
  <td colspan="11" align="center">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">MF-275'."\n".'MF-276</td>
 </tr>';
    push @html, $single_blank; 
    push @html, '
 <tr>
  <td colspan="30">Organic Rinse?</td>
  <td colspan="5" align="center">Yes</td>
  <td colspan="5" align="center">N/A</td>
  <td colspan="51" align="right">&nbsp;</td>
  <td colspan="9" align="right">MF-274</td>
 </tr>';
    push @html, $single_blank; 
    push @html, '
 <tr>
  <td colspan="30">Correct Capsule Size and Type?</td>
  <td colspan="5" align="center">Yes</th>
  <td colspan="5" align="center">No</th>
  <td colspan="65">&nbsp;</td>
 </tr>';
    push @html, $single_blank; 
#11 lines
    push @html, '<tbody class="QA">' . $single_blank;
    push @html, '
 <tr>
  <td colspan="30" class="heavy">QC Approval Only</td>
  <td colspan="28">&nbsp;</td>
  <td colspan="11" align="center" class="heavy">Date:&nbsp;_________</td>
  <td colspan="11" align="center" class="heavy">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center" class="heavy">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">&nbsp;</td>
 </tr>';
    push @html, $single_blank . '</tbody>'; 
    push @html, $single_blank;
    push @html, $single_blank;  
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;      
    push @html, $single_blank; 
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="30" class="heavy">Excipients Added</td>
  <td colspan="70">&nbsp;</td>
 </tr>';
    push @html, $single_blank;  
    push @html, '
 <tr>
  <td colspan="100" class="heavy" style="vertical-align: top"><b>Formulas:</b></td></tr>
 </tr>';  
    push @html, '
 <tr>
  <td colspan="50"><li style="list-style: disc inside"><b>Total Kg Added</b> = mg Per Unit Added X total capsules</li>
  <li style="list-style: disc inside"><b>mg Per Unit Added</b> = Total Kg Added &divide; total capsules</li></td>
  <td colspan="50">&nbsp;</td>
 </tr>';
    push @html, $single_blank;
#10 lines
    push @html, '
 <tr style="height:37px">
  <th colspan="9">Raw Material</td>
  <th colspan="18">Description</td>
  <th colspan="8">Lot #</td>
  <th colspan="10">Kg added</td>
  <th colspan="10">mg per unit</td>
  <th colspan="12">Purpose</td>
  <th colspan="6">Pre QC by</td>
  <th colspan="6">Added by</td>
  <th colspan="6">Blended by</td>
  <th colspan="6">Posted by</td>
  <th colspan="9">&nbsp;</td>
 </tr>';
    push @html, '
 <tr style="height:37px">
  <td colspan="9" align="center">__________</td>
  <td colspan="18" align="center">__________________________</td>
  <td colspan="8" align="center">__________</td>
  <td colspan="10" align="center">___________&nbsp;Kg</td>
  <td colspan="10" align="center">___________&nbsp;mg</td>
  <td colspan="12" align="center">______________________</td>
  <td colspan="6" align="center">________</td>
  <td colspan="6" align="center">________</td>
  <td colspan="6" align="center">________</td>
  <td colspan="6" align="center">________</td>
  <td colspan="9" align="right">QA-487</td> 
 </tr>';
    push @html, '
 <tr style="height:37px">
  <td colspan="9" align="center">__________</td>
  <td colspan="18" align="center">__________________________</td>
  <td colspan="8" align="center">__________</td>
  <td colspan="10" align="center">___________&nbsp;Kg</td>
  <td colspan="10" align="center">___________&nbsp;mg</td>
  <td colspan="12" align="center">______________________</td>
  <td colspan="6" align="center">________</td>
  <td colspan="6" align="center">________</td>
  <td colspan="6" align="center">________</td>
  <td colspan="6" align="center">________</td>
  <td colspan="9" align="right">QA-487</td> 
 </tr>';
    push @html, '
 <tr style="height:37px">
  <td colspan="9" align="center">__________</td>
  <td colspan="18" align="center">__________________________</td>
  <td colspan="8" align="center">__________</td>
  <td colspan="10" align="center">___________&nbsp;Kg</td>
  <td colspan="10" align="center">___________&nbsp;mg</td>
  <td colspan="12" align="center">______________________</td>
  <td colspan="6" align="center">________</td>
  <td colspan="6" align="center">________</td>
  <td colspan="6" align="center">________</td>
  <td colspan="6" align="center">________</td>
  <td colspan="9" align="right">QA-487</td> 
 </tr>';
 #8 lines
    push @html, $single_blank;      
    push @html, $single_blank;   
    push @html, $single_blank;  
    push @html, $single_blank;
    push @html, $single_blank;      
    push @html, $single_blank;   
    push @html, $single_blank;   
    push @html, $single_blank;  
    push @html, $single_blank;
    push @html, $single_blank;        
    push @html, $single_blank;   
    push @html, $single_blank;  
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;      
    push @html, $single_blank;   
    push @html, $single_blank;  
    push @html, $single_blank;      
    push @html, $main_tablef;
#24 lines    
    
   
# # encap_pg7
    $log->log( " > Printing 20 Capsule Weight Check Pg7...", 1 );
    push @html, $main_tableh;
    push @html, '
 <tr>
  <td colspan="100" class="heavy">20 Capsule Weight Check</td>
 </tr>';
    push @html, $single_blank;
    push @html, '
  <tr>
   <td colspan="22"><b>Target Capsule Weight</b><br>(with '.$cap_add.'&nbsp;mg capsule included)</td>
   <td colspan="13" class="heavy" align="center">' . $trg_weight . '&nbsp;mg</td>
   <td colspan="65">&nbsp;</td>
  </tr>';
    push @html, $single_blank;
    push @html, $single_blank;
#6 lines
    foreach my $count ( 1 .. 10 ) {
        push @html, '
 <tr style="height:46px">
  <td colspan="9">&nbsp;</td>
  <td colspan="13" style="text-align: center; vertical-align:middle;">&nbsp;Capsule&nbsp;' . $count . '&nbsp;</td>
  <td colspan="13" align="center">_____________&nbsp;mg</td>
  <td colspan="30">&nbsp;</td>
  <td colspan="13" style="text-align: center; vertical-align:middle;">&nbsp;Capsule&nbsp;' . ( $count + 10 ) . '&nbsp;</td>
  <td colspan="13" align="center">_____________&nbsp;mg</td>
  <td colspan="9">&nbsp;</td>
  </tr>';
    }
    push @html, $single_blank;
    push @html, '
 <tr style="height:46px">
  <td colspan="9">&nbsp;</td>
  <td colspan="13" class="heavy" align="center">20 Capsule<br>Average</td>
  <td colspan="13" align="center" class="heavy">____________&nbsp;mg</td>
  <td colspan="23">&nbsp;</td>
  <td colspan="11" align="center">Date:&nbsp;_________</td>
  <td colspan="11" align="center">Time:&nbsp;________<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">&nbsp;</td>
 </tr>';
    push @html, $single_blank;
    push @html, $single_blank;
#34 lines
    push @html, '
 <tr>
  <td colspan="100" class="heavy" style="vertical-align: top"><b>Formulas:</b></td>
 </tr>';  
    push @html, '
 <tr>
  <td colspan="50"><li style="list-style: disc inside"><b>lowest allowable mg</b> = 20 Capsule Average X 0.92</li>
  <li style="list-style: disc inside"><b>highest allowable mg</b> = 20 Capsule Average X 1.08</li></td>
  <td colspan="50">&nbsp;</td>
 </tr>';
    push @html, $single_blank;
    push @html, '
  <tr>
   <td colspan="100"><b>Capsule Parameters</b></td>
  </tr>';
    push @html, $single_blank;
    push @html, '
  <tr>
   <td colspan="30"><b>Capsule Weight Range</b><br>(with '.$cap_add.' mg capsule included)</td>
   <td colspan="10" class="heavy" align="center" style="vertical-align: bottom">__________&nbsp;mg</td>
   <td colspan="5" align="center">&nbsp;to&nbsp;</td>
   <td colspan="10" class="heavy" align="center" style="vertical-align: bottom">__________&nbsp;mg</td>
   <td colspan="14" align="right">(&plusmn;8&#37; of Target)<br>Conforms?</td>
   <td colspan="11" align="center">Yes</td>
   <td colspan="11" align="center">No</td>
   <td colspan="9">&nbsp;</td>
  </tr>';
    push @html, $single_blank;
    push @html, '
  <tr>
   <td colspan="30">Capsules locked and free from defects?</td>
   <td colspan="5" align="center">Yes</td>
   <td colspan="5" align="center">No</td>
   <td colspan="65">&nbsp;</td>
 </tr>';  
    push @html, $single_blank; 
    push @html, '
 <tr>
  <td colspan="30">Correct Capsule color?</td>
  <td colspan="5" align="center">Yes</th>
  <td colspan="5" align="center">No</th>
  <td colspan="65">&nbsp;</td>
 </tr>';
    push @html, $single_blank; 
    push @html, $single_blank;      
    push @html, '
 <tr>
  <td colspan="30" class="heavy">QC Approval Only</td>
  <td colspan="28">&nbsp;</td>
  <td colspan="11" align="center" class="heavy">Date:&nbsp;_________</td>
  <td colspan="11" align="center" class="heavy">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center" class="heavy">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">&nbsp;</td>
 </tr>';
    push @html, $single_blank;
    push @html, $main_tablef;
#15 lines


# # encap_pg8
    $log->log( " > Printing 10 Capsule Avg Weight Log pg8...", 1 );
    push @html, $main_tableh;
    push @html, '
 <tr>
  <td colspan="100" class="heavy">10 Capsule Average Weight Log</td>
 </tr>';
    push @html, '
 <tr>
  <td colspan="100"><li style="list-style: disc inside">After setting up the encapsulation fill weight parameters, fill in the Run Weight Log below every 15 minutes. Recorded weights are the average weight of 10 capsules. Weights include the weight of the capsules.</li>
  <li style="list-style: disc inside">When a drum is brought in Encapsulation Area, write and circle the drum number next to the appropriate time entry.</li></td>
 </tr>';
    push @html, $single_blank;
#5 lines
    push @html, '
  <tr>
   <td colspan="22"><b>Theoretical Capsule Target</b><br>(with '.$cap_add.'&nbsp;mg capsule included)</td>
   <td colspan="13" class="heavy" align="center">' . $trg_weight . '&nbsp;mg</td>
   <td colspan="65">&nbsp;</td>
  </tr>'; 
    push @html, $single_blank;
    push @html, '
 <tr style="height:46px">
  <th colspan="7" align="center">Date</th>
  <th colspan="9" align="center">Time</th>
  <th colspan="7" align="center">Avg.<br>Weight</th>
  <th colspan="10" align="center">Defects*<br><font size=-2>(Dents/Cracks/Tucks)</font></th>
  <th colspan="8" align="center">Locked?**</th>
  <th colspan="8" align="center">Initials</th>
  <th colspan="2" align="center">&nbsp;</td>
  <th colspan="7" align="center">Date</th>
  <th colspan="9" align="center">Time</th>
  <th colspan="7" align="center">Avg.<br>Weight</th>
  <th colspan="10" align="center">Defects&nbsp;*<br><font size=-2>(Dents/Cracks/Tucks)</font></th>
  <th colspan="8" align="center">Locked?&nbsp**</th>
  <th colspan="8" align="center">Initials</th>
 </tr>';
    foreach my $count ( 1 .. 12 ) {
        push @html, '
 <tr style="height:46px">
  <td colspan="7" align="center">__________</td>
  <td colspan="9" align="center">____:____<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="7" align="center">________&nbsp;mg&nbsp;</td>
  <td colspan="10" align="center">Yes&nbsp;/&nbsp;No</td>
  <td colspan="8" align="center">Yes&nbsp;/&nbsp;No</td>
  <td colspan="8" align="center">_________</td>
  <td colspan="2" align="center">&nbsp;</td>
  <td colspan="7" align="center">__________</td>
  <td colspan="9" align="center">____:____<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="7" align="center">________&nbsp;mg&nbsp;</td>
  <td colspan="10" align="center">Yes&nbsp;/&nbsp;No</td>
  <td colspan="8" align="center">Yes&nbsp;/&nbsp;No</td>
  <td colspan="8" align="center">_________</td>
</tr>';
    }
#41 lines
    push @html, $single_blank;
    push @html, '
 <tr style="height:46px">
   <td colspan="22" class="heavy">Average&nbsp;of&nbsp;Averages (this page)</td>
   <td colspan="13" align="center">_________&nbsp;mg</td>
   <td colspan="14" align="center">&nbsp;</td>
   <td colspan="9" class="heavy" align="center">Drums:</td>
   <td colspan="42" align="center">1&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;2&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;3&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;4&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;5&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;6&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;7&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;8&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;9&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;10&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;11&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;12</td>
  </tr>';
    push @html, $single_blank;
    push @html, '
  <tr>
   <td colspan="22"><b>Average&nbsp;Capsule Weight&nbsp;Range</b><br>(with '.$cap_add.' mg capsule included)</td>
   <td colspan="13" class="heavy" align="center" style="vertical-align:bottom">' . $lo_weight . '&nbsp;mg</td>
   <td colspan="10" align="center">and</td>
   <td colspan="13" class="heavy" align="center" style="vertical-align:bottom">' . $hi_weight . '&nbsp;mg</td>
   <td colspan="11" align="right">(&plusmn;5&#37; of Theoretical)<br>Conforms?</td>
   <td colspan="11" align="center">Yes</td>
   <td colspan="11" align="center">No</td>
   <td colspan="9">&nbsp;</td>
  </tr>';
    push @html, $single_blank; 
    push @html, '
 <tr>
  <td colspan="91"><b>&nbsp;&nbsp;*</b> If Capsule Defects are present, contact Quality Control for disposition.</td>
  <td colspan="9" align="center">&nbsp;</td>
 </tr>';
    push @html, '
 <tr>
  <td colspan="91"><b>**</b> If Capsules are not locked, contact Quality Control for disposition.</td>
  <td colspan="9" align="center">&nbsp;</td>
 </tr>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="58">If Average Capsule weight is outside of the range, contact Quality Control to address deviation.</td>
  <td colspan="11" class="heavy">QC Deviation:</td>
  <td colspan="11" align="center">#&nbsp;____________</td>
  <td colspan="11" align="center">N/A</td>
  <td colspan="9" align="right">QA-487</td>
 </tr>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="30" class="heavy">QC Approval Only</td>
  <td colspan="28">&nbsp;</td>
  <td colspan="11" align="center" class="heavy">Date:&nbsp;_________</td>
  <td colspan="11" align="center" class="heavy">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center" class="heavy">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">&nbsp;</td>
 </tr>';
    push @html, $single_blank;
    push @html, $main_tablef;
#15 lines


# # encap_pg9
    $log->log( " > Printing 10 Capsule Avg Weight Log Again pg9...", 1 );
    push @html, $main_tableh;
    push @html, '
 <tr>
  <td colspan="100" class="heavy">10 Capsule Average Weight Log</td>
 </tr>';
    push @html, '
 <tr>
  <td colspan="100"><li style="list-style: disc inside">After setting up the encapsulation fill weight parameters, fill in the Run Weight Log below every 15 minutes. Recorded weights are the average weight of 10 capsules. Weights include the weight of the capsules.</li>
  <li style="list-style: disc inside">When a drum is brought in Encapsulation Area, write and circle the drum number next to the appropriate time entry.</li></td>
 </tr>';
    push @html, $single_blank;
#5 lines
    push @html, '
  <tr>
   <td colspan="22"><b>Theoretical Capsule Target</b><br>(with '.$cap_add.'&nbsp;mg capsule included)</td>
   <td colspan="13" class="heavy" align="center">' . $trg_weight . '&nbsp;mg</td>
   <td colspan="65">&nbsp;</td>
  </tr>'; 
    push @html, $single_blank;
    push @html, '
 <tr style="height:46px">
  <th colspan="7" align="center">Date</th>
  <th colspan="9" align="center">Time</th>
  <th colspan="7" align="center">Avg.<br>Weight</th>
  <th colspan="10" align="center">Defects*<br><font size=-2>(Dents/Cracks/Tucks)</font></th>
  <th colspan="8" align="center">Locked?**</th>
  <th colspan="8" align="center">Initials</th>
  <th colspan="2" align="center">&nbsp;</td>
  <th colspan="7" align="center">Date</th>
  <th colspan="9" align="center">Time</th>
  <th colspan="7" align="center">Avg.<br>Weight</th>
  <th colspan="10" align="center">Defects&nbsp;*<br><font size=-2>(Dents/Cracks/Tucks)</font></th>
  <th colspan="8" align="center">Locked?&nbsp**</th>
  <th colspan="8" align="center">Initials</th>
 </tr>';
    foreach my $count ( 1 .. 12 ) {
        push @html, '
 <tr style="height:46px">
  <td colspan="7" align="center">__________</td>
  <td colspan="9" align="center">____:____<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="7" align="center">________&nbsp;mg&nbsp;</td>
  <td colspan="10" align="center">Yes&nbsp;/&nbsp;No</td>
  <td colspan="8" align="center">Yes&nbsp;/&nbsp;No</td>
  <td colspan="8" align="center">_________</td>
  <td colspan="2" align="center">&nbsp;</td>
  <td colspan="7" align="center">__________</td>
  <td colspan="9" align="center">____:____<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="7" align="center">________&nbsp;mg&nbsp;</td>
  <td colspan="10" align="center">Yes&nbsp;/&nbsp;No</td>
  <td colspan="8" align="center">Yes&nbsp;/&nbsp;No</td>
  <td colspan="8" align="center">_________</td>
</tr>';
    }
#41 lines
    push @html, $single_blank;
    push @html, '
 <tr style="height:46px">
   <td colspan="22" class="heavy">Average&nbsp;of&nbsp;Averages (this page)</td>
   <td colspan="13" align="center">_________&nbsp;mg</td>
   <td colspan="14" align="center">&nbsp;</td>
   <td colspan="9" class="heavy" align="center">Drums:</td>
   <td colspan="42" align="center">1&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;2&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;3&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;4&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;5&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;6&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;7&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;8&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;9&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;10&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;11&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;12</td>
  </tr>';
    push @html, $single_blank;
    push @html, '
  <tr>
   <td colspan="22"><b>Average&nbsp;Capsule Weight&nbsp;Range</b><br>(with '.$cap_add.' mg capsule included)</td>
   <td colspan="13" class="heavy" align="center" style="vertical-align:bottom">' . $lo_weight . '&nbsp;mg</td>
   <td colspan="10" align="center">and</td>
   <td colspan="13" class="heavy" align="center" style="vertical-align:bottom">' . $hi_weight . '&nbsp;mg</td>
   <td colspan="11" align="right">(&plusmn;5&#37; of Theoretical)<br>Conforms?</td>
   <td colspan="11" align="center">Yes</td>
   <td colspan="11" align="center">No</td>
   <td colspan="9">&nbsp;</td>
  </tr>';
    push @html, $single_blank; 
    push @html, '
 <tr>
  <td colspan="91"><b>&nbsp;&nbsp;*</b> If Capsule Defects are present, contact Quality Control for disposition.</td>
  <td colspan="9" align="center">&nbsp;</td>
 </tr>';
    push @html, '
 <tr>
  <td colspan="91"><b>**</b> If Capsules are not locked, contact Quality Control for disposition.</td>
  <td colspan="9" align="center">&nbsp;</td>
 </tr>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="58">If Average Capsule weight is outside of the range, contact Quality Control to address deviation.</td>
  <td colspan="11" class="heavy">QC Deviation:</td>
  <td colspan="11" align="center">#&nbsp;____________</td>
  <td colspan="11" align="center">N/A</td>
  <td colspan="9" align="right">QA-487</td>
 </tr>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="30" class="heavy">QC Approval Only</td>
  <td colspan="28">&nbsp;</td>
  <td colspan="11" align="center" class="heavy">Date:&nbsp;_________</td>
  <td colspan="11" align="center" class="heavy">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center" class="heavy">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">&nbsp;</td>
 </tr>';
    push @html, $single_blank;
    push @html, $main_tablef;
#15 lines


# # encap_pg10
    $log->log( " > Printing Capsule Yield Log Pg10...", 1 );
    push @html, $main_tableh;
    push @html, '
  <tr>
   <td colspan="100" class="heavy">Capsule Yield Log</td>
 </tr>';
    push @html, '
 <tr>
  <td colspan="100"><li style="list-style: disc inside">Once the lot is encapsulated, record the Production Yield in the spaces below.</li></td>
 </tr>';
    push @html, $single_blank;  
#5 lines
    foreach my $count ( 1 .. 7 ) {
        push @html, '
 <tr style="height:46px">
  <td colspan="9">&nbsp;</td>
  <td colspan="12" align="center">Container #' . $count . '</td>
  <td colspan="12" align="center">____________&nbsp;Kg</td>
  <td colspan="5" align="center">&nbsp;</td>
  <td colspan="12" align="center">Container #' . ($count+7) . '</td>
  <td colspan="12" align="center">____________&nbsp;Kg</td>
  <td colspan="5" align="center">&nbsp;</td>
  <td colspan="12" align="center">Container #' . ($count+14) . '</td>
  <td colspan="12" align="center">____________&nbsp;Kg</td>
  <td colspan="9">&nbsp;</td>
 </tr>';
    }

    push @html, $single_blank;
#    push @html, '
# <tr class="double">
#  <td colspan="2">&nbsp;</td>
#  <td colspan="2" align="center">Avg.&nbsp;Capsule&nbsp;Weight</td>
#  <td colspan="2" align="center">_________&nbsp;mg</td>
#  <td colspan="1" align="center">&minus;</td>
#  <td colspan="2" align="center">Empty Capsule Weight</td>
#  <td colspan="2" align="center"><u>&nbsp;&nbsp;&nbsp;'. $cap_add .'&nbsp;&nbsp;&nbsp;</u>&nbsp;mg</td>
#  <td colspan="2" align="center">&nbsp;</td>  
#  <td colspan="2" align="center"><b>= Avg. Fill Weight</b></td>
#  <td colspan="3">_________&nbsp;mg</td>
# </tr>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="9">&nbsp;</td>
  <td colspan="12" align="center">Total Weight</td>
  <td colspan="12" align="center">_________&nbsp;Kg</td>
  <td colspan="10" align="center">&divide;</td>
  <td colspan="12" align="center">Weight&nbsp;of&nbsp;Avg.&nbsp;Capsule</td>'.
# <td colspan="2" align="center">Avg.&nbsp;Fill&nbsp;Weight</td>
  '<td colspan="12" align="center">_________&nbsp;mg</td>
  <td colspan="12" align="center">X 1,000,000</td>  
  <td colspan="12" align="center"><b>= Total Capsules</b></td>
  <td colspan="9">_________</td>
 </tr>';
    push @html, $single_blank;
#    push @html, '
# <tr class="single">
#  <td colspan="2" class="heavy" align="right">Example:</td>
#  <td colspan="2" align="center">Total Weight</td>
#  <td colspan="2" align="center"><u><i>&nbsp;&nbsp;' . $runsize . '&nbsp;&nbsp;</i></u>&nbsp;Kg</td>
#  <td colspan="1" align="center">&divide;</td>
#  <td colspan="2" align="center">Avg.&nbsp;Fill&nbsp;Weight</td>
#  <td colspan="2" align="center"><u><i>&nbsp;&nbsp;&nbsp;' . $fillweight . '&nbsp;&nbsp;&nbsp;</i></u>&nbsp;mg</td>
#  <td colspan="2" align="center">X 1,000,000</td>   
#  <td colspan="2" align="center"><b>= Total Capsules</b></td>
#  <td colspan="3"><u><i>&nbsp;&nbsp;&nbsp;' . sprintf("%0.f",$runsize/$fillweight*1000000) . '&nbsp;&nbsp;&nbsp;</i></u></td>
# </tr>';
    push @html, '
 <tr>
  <td colspan="9" class="heavy" align="center">Example:</td>
  <td colspan="12" align="center">Total Weight</td>
  <td colspan="12" align="center"><u><i>&nbsp;&nbsp;&nbsp;500.25&nbsp;&nbsp;&nbsp;&nbsp;</i></u>&nbsp;Kg</td>
  <td colspan="10" align="center">&divide;</td>
  <td colspan="12" align="center">Weight&nbsp;of&nbsp;Avg.&nbsp;Capsule</td>'.
# <td colspan="2" align="center">Avg.&nbsp;Fill&nbsp;Weight</td>
  ' <td colspan="12" align="center"><u><i>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;825&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</i></u>&nbsp;mg</td>
  <td colspan="12" align="center">X 1,000,000</td>   
  <td colspan="12" align="center"><b>= Total Capsules</b></td>
  <td colspan="9"><u><i>&nbsp;&nbsp;606,364&nbsp;&nbsp;</i></u></td>
 </tr>';
    push @html, $single_blank;
    push @html, '
  <tr>
   <td colspan="21" class="heavy">Theoretical Yield:</td>
   <td colspan="12" align="center" class="heavy">' . $theo_yield . ' capsules</td>
   <td colspan="67">&nbsp;</td>
 </tr>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="21" class="heavy">Theoretical Yield<br>must be between</td>
  <td colspan="12" align="center" class="heavy" style="vertical-align: bottom">' . $lo_yield . ' capsules</td>
  <td colspan="10" align="center">and</td>
  <td colspan="12" align="center" class="heavy" style="vertical-align: bottom">' . $hi_yield . ' capsules</td>
  <td colspan="12" align="right">(&plusmn;3&#37; of Theoretical)<br>Conforms?</td>
  <td colspan="12" align="center">Yes</td>
  <td colspan="12" align="center">No&nbsp;<i>(requires&nbsp;reconciliation)</i></td>
  <td colspan="9">&nbsp;</td>
 </tr>';
    push @html, $single_blank;
#16 lines
    push @html, '
 <tr>
  <td colspan="100">If Yield is outside of the range, place product on hold and contact Production Manager and Quality for Yield Reconciliation.</td>
 </tr>';
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, '
  <tr>
   <td colspan="30">Encapsulation started by</td>
   <td colspan="28">&nbsp;</td>
   <td colspan="11" align="center">Date:&nbsp;_________</td>
   <td colspan="11" align="center">Time:&nbsp;________<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
   <td colspan="11" align="center">Initials:&nbsp;_________</td>
   <td colspan="9">&nbsp;</td>
 </tr>';
    push @html, $single_blank;
    push @html, '
  <tr>
   <td colspan="30">Encapsulation finished by</td>
   <td colspan="28">&nbsp;</td>
   <td colspan="11" align="center">Date:&nbsp;_________</td>
   <td colspan="11" align="center">Time:&nbsp;________<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
   <td colspan="11" align="center">Initials:&nbsp;_________</td>
   <td colspan="9">&nbsp;</td>
 </tr>';
    push @html, $single_blank;
    push @html, '
  <tr>
   <td colspan="30">Hand-sort needed?</td>
   <td colspan="5" align="center">Yes</td>
   <td colspan="5" align="center">No</td>
   <td colspan="18">&nbsp;</td>
   <td colspan="11" align="center">Date:&nbsp;_________</td>
   <td colspan="11" align="center">Time:&nbsp;________<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
   <td colspan="11" align="center">Initials:&nbsp;_________</td>
   <td colspan="9">&nbsp;</td>
 </tr>';
    push @html, $single_blank;
    push @html, '
  <tr>
   <td colspan="30">Hand-sort Loss</td>
   <td colspan="10" align="center">_________&nbsp;Kg</td>
   <td colspan="18">&nbsp;</td>
   <td colspan="11" align="center">Date:&nbsp;_________</td>
   <td colspan="11" align="center">Time:&nbsp;________<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
   <td colspan="11" align="center">Initials:&nbsp;_________</td>
   <td colspan="9">&nbsp;</td>
 </tr>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="30">Other Losses/Leftover Powder</td>
  <td colspan="10" align="center">_________&nbsp;Kg</td>
  <td colspan="18">&nbsp;</td>
  <td colspan="11" align="center">Date:&nbsp;_________</td>
  <td colspan="11" align="center">Time:&nbsp;________<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center">Initials:&nbsp;_________</td> 
  <td colspan="9">&nbsp;</td>
 </tr>';
    push @html, $single_blank;
    push @html, '
  <tr>
   <td colspan="30" class="heavy">QC Batch Yield Reconciliation?</td>
   <td colspan="5" align="center" class="heavy">Yes</td>
   <td colspan="5" align="center" class="heavy">N/A</td>
   <td colspan="18">&nbsp;</td>
   <td colspan="11" align="center" class="heavy">Date:&nbsp;_________</td>
   <td colspan="11" align="center" class="heavy">Time:&nbsp;________<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
   <td colspan="11" align="center" class="heavy">Initials:&nbsp;_________</td>
   <td colspan="9">&nbsp;</td>
  </tr>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="30" class="heavy">QC Approval Only</td>
  <td colspan="28">&nbsp;</td>
  <td colspan="11" align="center" class="heavy">Date:&nbsp;_________</td>
  <td colspan="11" align="center" class="heavy">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center" class="heavy">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">&nbsp;</td>
 </tr>';
    push @html, $single_blank;
    push @html, $main_tablef;
#14 lines


# # encap_pg11
    $log->log( " > Printing Post-Encapsulation Pg11...", 1 );
    push @html, $main_tableh;
    push @html, '
 <tr>
   <td colspan="100" class="heavy">Polishing and Dedusting Check</td>
 </tr>';
    push @html, $single_blank; 
    push @html, '
 <tr>
  <td colspan="30">Room #_______ Cleaned</td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="23">&nbsp;</td>
  <td colspan="11" align="center">Date:&nbsp;_________</td>
  <td colspan="11" align="center">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">MF-275'."\n".'MF-276</td>
 </tr>';
    push @html, $single_blank; 
    push @html, '
 <tr>
  <td colspan="30">Organic Rinse?</td>
  <td colspan="5" align="center">Yes</td>
  <td colspan="5" align="center">N/A</td>
  <td colspan="51" align="right">&nbsp;</td>
  <td colspan="9" align="right">MF-274</td>
 </tr>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="30">Product Dedusted</td>
  <td colspan="5" align="center">'. $check_box .'</td>
  <td colspan="23">&nbsp;</td>
  <td colspan="11" align="center">Date:&nbsp;_________</td>
  <td colspan="11" align="center">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">&nbsp;</td>
 </tr>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="30">Waste Caps/Powder _______ Kg</td>
  <td colspan="28">&nbsp;</td>
  <td colspan="11" align="center">Date:&nbsp;_________</td>
  <td colspan="11" align="center">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center">Initials:&nbsp;_________</td>
  <td colspan="9">&nbsp;</td>
 </tr>';
    push @html, $single_blank; 
#11 lines
    push @html, '<tbody class="QA">' . $single_blank;
    push @html, '
 <tr>
  <td colspan="30" class="heavy">QC Approval Only</td>
  <td colspan="28">&nbsp;</td>
  <td colspan="11" align="center" class="heavy">Date:&nbsp;_________</td>
  <td colspan="11" align="center" class="heavy">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center" class="heavy">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">&nbsp;</td>
 </tr>';
    push @html, $single_blank . '</tbody>'; 
    push @html, $single_blank;  
    push @html, $single_blank; 
    push @html, $single_blank; 
    push @html, $single_blank; 
    push @html, $single_blank; 
    push @html, $single_blank; 
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank; 
#12 lines
    push @html, '
 <tr>
  <td colspan="100" align="left" class="heavy">Portable Equipment Used</td>
 </tr>';
    push @html, $single_blank; 
    push @html, '
 <tr>
  <td colspan="30">Equipment&nbsp;Used (circle&nbsp;one):</td>
  <td colspan="15" align="center" class="heavy">Colton Mill</td>
  <td colspan="15" align="center" class="heavy">Hammer Mill</td>
  <td colspan="15" align="center" class="heavy">Other</td>
  <td colspan="25">&nbsp;</td>
 </tr>';
    push @html, $single_blank; 
    push @html, '
 <tr>
  <td colspan="30">Room #________ Cleaned</td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="23">&nbsp;</td>
  <td colspan="11" align="center">Date:&nbsp;_________</td>
  <td colspan="11" align="center">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">MF-275'."\n".'MF-276</td>
 </tr>';
    push @html, $single_blank; 
    push @html, '
 <tr>
  <td colspan="30">Equipment Cleaned</td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="23">&nbsp;</td>
  <td colspan="11" align="center">Date:&nbsp;_________</td>
  <td colspan="11" align="center">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">MF-274</td>
 </tr>';  
    push @html, $single_blank; 
    push @html, '
 <tr>
  <td colspan="30">Organic Rinse?</td>
  <td colspan="5" align="center">Yes</td>
  <td colspan="5" align="center">N/A</td>
  <td colspan="51" align="right">&nbsp;</td>
  <td colspan="9" align="right">MF-274</td>
</tr>';
    push @html, $single_blank; 
    push @html, '
 <tr>
  <td colspan="30">Incoming Weight</td>
  <td colspan="5" align="center">_________&nbsp;Kg</td>
  <td colspan="23">&nbsp;</td>
  <td colspan="11" align="center">Date:&nbsp;_________</td>
  <td colspan="11" align="center">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center">Initials:&nbsp;_________</td>
  <td colspan="9" align="right"></td>
 </tr>';  
    push @html, $single_blank; 
    push @html, '
 <tr>
  <td colspan="30">Outgoing Weight</td>
  <td colspan="5" align="center">_________&nbsp;Kg</td>
  <td colspan="23">&nbsp;</td>
  <td colspan="11" align="center">Date:&nbsp;_________</td>
  <td colspan="11" align="center">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center">Initials:&nbsp;_________</td>
  <td colspan="9" align="right"></td>
 </tr>';
 #14 lines  
    push @html, $single_blank; 
    push @html, '<tbody class="QA">' . $single_blank;
    push @html, '
 <tr>
  <td colspan="30" class="heavy">QC Approval Only</td>
  <td colspan="28">&nbsp;</td>
  <td colspan="11" align="center" class="heavy">Date:&nbsp;_________</td>
  <td colspan="11" align="center" class="heavy">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center" class="heavy">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">&nbsp;</td>
 </tr>';
    push @html, $single_blank . '</tbody>'; 
    push @html, $single_blank;
    push @html, $single_blank; 
    push @html, $single_blank; 
    push @html, $single_blank; 
    push @html, $single_blank;     
    push @html, '</table><div class="footer"></div>';
#    push @html, $main_tablef;
    return ( 1, @html );
}



sub mmr_tab {
#####
    #
    # Section 2 of MMR - Tableting 
    #
#####
    my @args      = @_;
    my $form_obj  = shift @args;
    my $ppro_obj  = shift @args;
    my $spec_obj  = shift @args;
    my $batch_obj = shift @args;
    my $runsize   = shift @args;
    my @html      = @args;

    if ( !$ppro_obj->{'PRWeight'} ) {
        $log->log( "!! Unable to locate Fill Weight in ProcessPro!", 1 );
        return 0;
    }

    my $lo_range   = sprintf "%.3f", $runsize * .99;
    my $hi_range   = sprintf "%.3f", $runsize * 1.01;
    my $fillweight = sprintf "%.0f", $ppro_obj->{'PRWeight'} * 1000;
    my $lo_weight  = sprintf "%.0f", $fillweight * 0.95;
    my $hi_weight  = sprintf "%.0f", $fillweight * 1.05;
    my $theo_yield = sprintf "%.0f", $runsize / $fillweight * 1000000;
    my $lo_yield   = sprintf "%.0f", $theo_yield * .97;
    my $hi_yield   = sprintf "%.0f", $theo_yield * 1.03;
    $hi_yield =~ s/(\d{1,3}?)(?=(\d{3})+$)/$1,/gm;
    $fillweight =~ s/(\d{1,3}?)(?=(\d{3})+$)/$1,/gm;
    $theo_yield =~ s/(\d{1,3}?)(?=(\d{3})+$)/$1,/gm;
    $lo_weight =~ s/(\d{1,3}?)(?=(\d{3})+$)/$1,/gm;
    $hi_weight =~ s/(\d{1,3}?)(?=(\d{3})+$)/$1,/gm;
    $lo_yield =~ s/(\d{1,3}?)(?=(\d{3})+$)/$1,/gm;


# # tab_pg6
    $log->log( " > Printing Tableting Clean Check Pg6...", 1 );
    push @html, $main_tableh;
    push @html, '
 <tr>
  <td colspan="100" align="left" class="heavy">Tableting Area Clean Check</td>
 </tr>';
    push @html, $single_blank;
    push @html, ' 
 <tr>
  <td colspan="30">Equipment #_______ Cleaned</td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="23">&nbsp;</td>
  <td colspan="11" align="center">Date:&nbsp;_________</td>
  <td colspan="11" align="center">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">MF-285</td>
 </tr>';
    push @html, $single_blank; 
    push @html, '
 <tr>
  <td colspan="30">Room #_______ Cleaned</td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="23">&nbsp;</td>
  <td colspan="11" align="center">Date:&nbsp;_________</td>
  <td colspan="11" align="center">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">MF-285</td>
 </tr>';
    push @html, $single_blank; 
    push @html, '
 <tr>
  <td colspan="30">Organic Rinse?</td>
  <td colspan="5" align="center">Yes</td>
  <td colspan="5" align="center">N/A</td>
  <td colspan="51" align="right">&nbsp;</td>
  <td colspan="9" align="right">MF-274</td>
 </tr>';
    push @html, $single_blank; 
    push @html, '
 <tr>
  <td colspan="30">Correct Punch Size and Type?</td>
  <td colspan="5" align="center">Yes</th>
  <td colspan="5" align="center">No</th>
  <td colspan="65">&nbsp;</td>
 </tr>';
    push @html, $single_blank; 
#11 lines
    push @html, '<tbody class="QA">' . $single_blank;
    push @html, '
 <tr>
  <td colspan="30" class="heavy">QC Approval Only</td>
  <td colspan="28">&nbsp;</td>
  <td colspan="11" align="center" class="heavy">Date:&nbsp;_________</td>
  <td colspan="11" align="center" class="heavy">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center" class="heavy">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">&nbsp;</td>
 </tr>';
    push @html, $single_blank . '</tbody>'; 
    push @html, $single_blank; 
    push @html, $single_blank;  
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;      
    push @html, $single_blank; 
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="30" class="heavy">Excipients Added</td>
  <td colspan="70">&nbsp;</td>
 </tr>';
    push @html, $single_blank;  
    push @html, '
 <tr>
  <td colspan="100" class="heavy" style="vertical-align: top"><b>Formulas:</b></td></tr>
 </tr>';  
    push @html, '
 <tr>
  <td colspan="50"><li style="list-style: disc inside"><b>Total Kg Added</b> = mg Per Unit Added X total capsules</li>
  <li style="list-style: disc inside"><b>mg Per Unit Added</b> = Total Kg Added &divide; total capsules</li></td>
  <td colspan="50">&nbsp;</td>
 </tr>';
    push @html, $single_blank;
#10 lines
    push @html, '
 <tr style="height:37px">
  <th colspan="9">Raw Material</td>
  <th colspan="18">Description</td>
  <th colspan="8">Lot #</td>
  <th colspan="10">Kg added</td>
  <th colspan="10">mg per unit</td>
  <th colspan="12">Purpose</td>
  <th colspan="6">Pre QC by</td>
  <th colspan="6">Added by</td>
  <th colspan="6">Blended by</td>
  <th colspan="6">Posted by</td>
  <th colspan="9">&nbsp;</td>
 </tr>';
    push @html, '
 <tr style="height:37px">
  <td colspan="9" align="center">__________</td>
  <td colspan="18" align="center">__________________________</td>
  <td colspan="8" align="center">__________</td>
  <td colspan="10" align="center">___________&nbsp;Kg</td>
  <td colspan="10" align="center">___________&nbsp;mg</td>
  <td colspan="12" align="center">______________________</td>
  <td colspan="6" align="center">________</td>
  <td colspan="6" align="center">________</td>
  <td colspan="6" align="center">________</td>
  <td colspan="6" align="center">________</td>
  <td colspan="9" align="right">QA-487</td> 
 </tr>';
    push @html, '
 <tr style="height:37px">
  <td colspan="9" align="center">__________</td>
  <td colspan="18" align="center">__________________________</td>
  <td colspan="8" align="center">__________</td>
  <td colspan="10" align="center">___________&nbsp;Kg</td>
  <td colspan="10" align="center">___________&nbsp;mg</td>
  <td colspan="12" align="center">______________________</td>
  <td colspan="6" align="center">________</td>
  <td colspan="6" align="center">________</td>
  <td colspan="6" align="center">________</td>
  <td colspan="6" align="center">________</td>
  <td colspan="9" align="right">QA-487</td> 
 </tr>';
    push @html, '
 <tr style="height:37px">
  <td colspan="9" align="center">__________</td>
  <td colspan="18" align="center">__________________________</td>
  <td colspan="8" align="center">__________</td>
  <td colspan="10" align="center">___________&nbsp;Kg</td>
  <td colspan="10" align="center">___________&nbsp;mg</td>
  <td colspan="12" align="center">______________________</td>
  <td colspan="6" align="center">________</td>
  <td colspan="6" align="center">________</td>
  <td colspan="6" align="center">________</td>
  <td colspan="6" align="center">________</td>
  <td colspan="9" align="right">QA-487</td> 
 </tr>';
 #8 lines
    push @html, $single_blank;      
    push @html, $single_blank;   
    push @html, $single_blank;  
    push @html, $single_blank;
    push @html, $single_blank;      
    push @html, $single_blank;   
    push @html, $single_blank;  
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;      
    push @html, $single_blank;   
    push @html, $single_blank;  
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;      
    push @html, $single_blank;   
    push @html, $single_blank;  
    push @html, $single_blank;      
    push @html, $main_tablef;
#24 lines 
   
   

# # tab_pg7
    $log->log( " > Printing 20 Tablet Weight Check Pg7...", 1 );
    push @html, $main_tableh;
    push @html, '
 <tr>
  <td colspan="100" class="heavy">20 Tablet Weight Check</td>
 </tr>';
    push @html, $single_blank;
    push @html, '
  <tr>
   <td colspan="22"><b>Target Tablet Weight</b></td>
   <td colspan="13" class="heavy" align="center">' . $fillweight . '&nbsp;mg</td>
   <td colspan="65">&nbsp;</td>
  </tr>';
    push @html, $single_blank;
    push @html, $single_blank;
#6 lines
    foreach my $count ( 1 .. 10 ) {
        push @html, '
 <tr style="height:46px">
  <td colspan="9">&nbsp;</td>
  <td colspan="13" style="text-align: center; vertical-align:middle;">&nbsp;Tablet&nbsp;' . $count . '&nbsp;</td>
  <td colspan="13" align="center">_____________&nbsp;mg</td>
  <td colspan="30">&nbsp;</td>
  <td colspan="13" style="text-align: center; vertical-align:middle;">&nbsp;Tablet&nbsp;' . ( $count + 10 ) . '&nbsp;</td>
  <td colspan="13" align="center">_____________&nbsp;mg</td>
  <td colspan="9">&nbsp;</td>
  </tr>';
    }
    push @html, $single_blank;
    push @html, '
 <tr style="height:46px">
  <td colspan="9">&nbsp;</td>
  <td colspan="13" class="heavy" align="center">20 Tablet<br>Average</td>
  <td colspan="13" align="center" class="heavy">____________&nbsp;mg</td>
  <td colspan="23">&nbsp;</td>
  <td colspan="11" align="center">Date:&nbsp;_________</td>
  <td colspan="11" align="center">Time:&nbsp;________<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">&nbsp;</td>
 </tr>';
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="100" class="heavy" style="vertical-align: top"><b>Formulas:</b></td>
 </tr>';
    push @html, '
 <tr>  
  <td colspan="50"><li style="list-style: disc inside"><b>lowest allowable mg</b> = 20 Tablet Average X 0.92</li>
  <li style="list-style: disc inside"><b>highest single mg</b> = 20 Tablet Average X 1.08</li></td>
  <td colspan="50">&nbsp;</td>
 </tr>';
#17 lines
    push @html, '
  <tr>
   <td colspan="100"><b>Tablet Parameters</b></td>
  </tr>';
    push @html, $single_blank;
    push @html, '
  <tr>
   <td colspan="30">Tablet Weight Range</td>
   <td colspan="10" class="heavy" align="center" style="vertical-align: bottom">_________&nbsp;mg</td>
   <td colspan="5" align="center">&nbsp;to&nbsp;</td>
   <td colspan="10" class="heavy" align="center" style="vertical-align: bottom">_________&nbsp;mg</td>
   <td colspan="14" align="right">(&plusmn;8&#37;)<br>Conforms?</td>
   <td colspan="11" align="center">Yes</td>
   <td colspan="11" align="center">No</td>
   <td colspan="9">&nbsp;</td>
  </tr>';
    push @html, $single_blank;
    push @html, '
  <tr>
   <td colspan="20">Tablet Thickness Target</td>
   <td colspan="10" align="center">_________&nbsp;mm</td>
   <td colspan="10" class="heavy" align="center" style="vertical-align: bottom">(_________&nbsp;mm</td>
   <td colspan="5" align="center">&nbsp;to&nbsp;</td>
   <td colspan="10" class="heavy" align="center" style="vertical-align: bottom">_________&nbsp;mm)</td>
   <td colspan="14" align="right">(&plusmn;5&#37;)<br>Conforms?</td>
   <td colspan="11" align="center">Yes</td>
   <td colspan="11" align="center">No</td>
   <td colspan="9">&nbsp;</td>
  </tr>';
    push @html, $single_blank;
    push @html, '
  <tr>
   <td colspan="20">Tablet Hardness Target</td>
   <td colspan="10" align="center">_________&nbsp;Kp</td>
   <td colspan="10" class="heavy" align="center" style="vertical-align: bottom">(_________&nbsp;Kp</td>
   <td colspan="5" align="center">&nbsp;to&nbsp;</td>
   <td colspan="10" class="heavy" align="center" style="vertical-align: bottom">_________&nbsp;Kp)</td>
   <td colspan="14" align="right">(&plusmn;10&#37;)<br>Conforms?</td>
   <td colspan="11" align="center">Yes</td>
   <td colspan="11" align="center">No</td>
   <td colspan="9">&nbsp;</td>
  </tr>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="58">If the Tablet parameters are outside of the range,<br>place product on hold and contact Production Manager and Quality.</td>
  <td colspan="11" class="heavy">QC Deviation:</td>
  <td colspan="11" align="center">#&nbsp;____________</td>
  <td colspan="11" align="center">N/A</td>
  <td colspan="9" align="right">QA-487</td>
 </tr>';
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="30" class="heavy">QC Approval Only</td>
  <td colspan="28">&nbsp;</td>
  <td colspan="11" align="center" class="heavy">Date:&nbsp;_________</td>
  <td colspan="11" align="center" class="heavy">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center" class="heavy">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">&nbsp;</td>
 </tr>';
    push @html, $single_blank;
    push @html, $main_tablef;
#14 lines



# # tab_pg8
    $log->log( " > Printing 10 Tablet Avg Weight Log Pg8...", 1 );
    push @html, $main_tableh;
    push @html, '
 <tr>
  <td colspan="100" class="heavy">10 Tablet Average Weight Log</td>
 </tr>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="100"><li style="list-style: disc inside">After setting up the Tableting weight parameters, fill in the Run Weight Log below every 15 minutes. Recorded weights are the average weight of 10 tablets.</li>
  <li style="list-style: disc inside">When a drum is brought in Encapsulation Area, write and circle the drum number next to the appropriate time entry.</li></td>
 </tr>';
    push @html, $single_blank;
#5 lines
    push @html, '
  <tr>
   <td colspan="22"><b>Theoretical Tablet Target</b><br></td>
   <td colspan="13" class="heavy" align="center">' . $fillweight . '&nbsp;mg</td>
   <td colspan="65">&nbsp;</td>
  </tr>';
    push @html, $single_blank;
     push @html, '
 <tr style="height:46px">
  <th colspan="9" align="center">Date</th>
  <th colspan="9" align="center">Time</th>
  <th colspan="9" align="center">Avg.<br>Weight</th>
  <th colspan="10" align="center">Defects*<br><font size=-2>(Dents/Cracks/Tucks)</font></th>
  <th colspan="9" align="center">Initials</th>
  <th colspan="8" align="center">&nbsp;</td>
  <th colspan="9" align="center">Date</th>
  <th colspan="9" align="center">Time</th>
  <th colspan="9" align="center">Avg.<br>Weight</th>
  <th colspan="10" align="center">Defects&nbsp;*<br><font size=-2>(Dents/Cracks/Tucks)</font></th>
  <th colspan="9" align="center">Initials</th>
 </tr>';
    foreach my $count ( 1 .. 12 ) {
        push @html, '
 <tr style="height:46px">
  <td colspan="9" align="center">__________</td>
  <td colspan="9" align="center">____:____<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="9" align="center">________&nbsp;mg&nbsp;</td>
  <td colspan="10" align="center">Yes&nbsp;/&nbsp;No</td>
  <td colspan="9" align="center">_________</td>
  <td colspan="8" align="center">&nbsp;</td>
  <td colspan="9" align="center">__________</td>
  <td colspan="9" align="center">____:____<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="9" align="center">________&nbsp;mg&nbsp;</td>
  <td colspan="10" align="center">Yes&nbsp;/&nbsp;No</td>
  <td colspan="9" align="center">_________</td>
 </tr>';
    }
    push @html, $single_blank;
#41 lines
    push @html, '
 <tr style="height:46px">
   <td colspan="22" class="heavy">Average&nbsp;of&nbsp;Averages (this page)</td>
   <td colspan="13" align="center">_________&nbsp;mg</td>
   <td colspan="14" align="center">&nbsp;</td>
   <td colspan="9" class="heavy" align="center">Drums:</td>
   <td colspan="42" align="center">1&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;2&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;3&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;4&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;5&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;6&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;7&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;8&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;9&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;10&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;11&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;12</td>
  </tr>';
    push @html, $single_blank;
    push @html, '
  <tr>
   <td colspan="22"><b>Average&nbsp;Tablet Weight&nbsp;Range</b></td>
   <td colspan="13" class="heavy" align="center" style="vertical-align:bottom">' . $lo_weight . '&nbsp;mg</td>
   <td colspan="10" align="center">and</td>
   <td colspan="13" class="heavy" align="center" style="vertical-align:bottom">' . $hi_weight . '&nbsp;mg</td>
   <td colspan="11" align="right">(&plusmn;5&#37; of Theoretical)<br>Conforms?</td>
   <td colspan="11" align="center">Yes</td>
   <td colspan="11" align="center">No</td>
   <td colspan="9">&nbsp;</td>
  </tr>';
    push @html, $single_blank; 
    push @html, '
 <tr>
  <td colspan="91"><b>&nbsp;&nbsp;*</b> If Capsule Defects are present, contact Quality Control for disposition.</td>
  <td colspan="9" align="center">&nbsp;</td>
 </tr>';
    push @html, '
 <tr>
  <td colspan="58">If Average Tablet weight is outside of the range, contact Quality Control to address deviation.</td>
  <td colspan="11" class="heavy">QC Deviation:</td>
  <td colspan="11" align="center">#&nbsp;____________</td>
  <td colspan="11" align="center">N/A</td>
  <td colspan="9" align="right">QA-487</td>
 </tr>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="30" class="heavy">QC Approval Only</td>
  <td colspan="28">&nbsp;</td>
  <td colspan="11" align="center" class="heavy">Date:&nbsp;_________</td>
  <td colspan="11" align="center" class="heavy">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center" class="heavy">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">&nbsp;</td>
 </tr>';
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $main_tablef;
#12 lines




# # tab_pg9
    $log->log( " > Printing 10 Tablet Avg Weight Log Pg9...", 1 );
    push @html, $main_tableh;
    push @html, '
 <tr>
  <td colspan="100" class="heavy">10 Tablet Average Weight Log</td>
 </tr>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="100"><li style="list-style: disc inside">After setting up the Tableting weight parameters, fill in the Run Weight Log below every 15 minutes. Recorded weights are the average weight of 10 tablets.</li>
  <li style="list-style: disc inside">When a drum is brought in Encapsulation Area, write and circle the drum number next to the appropriate time entry.</li></td>
 </tr>';
    push @html, $single_blank;
#5 lines
    push @html, '
  <tr>
   <td colspan="22"><b>Theoretical Tablet Target</b><br></td>
   <td colspan="13" class="heavy" align="center">' . $fillweight . '&nbsp;mg</td>
   <td colspan="65">&nbsp;</td>
  </tr>';
    push @html, $single_blank;
     push @html, '
 <tr style="height:46px">
  <th colspan="9" align="center">Date</th>
  <th colspan="9" align="center">Time</th>
  <th colspan="9" align="center">Avg.<br>Weight</th>
  <th colspan="10" align="center">Defects*<br><font size=-2>(Dents/Cracks/Tucks)</font></th>
  <th colspan="9" align="center">Initials</th>
  <th colspan="8" align="center">&nbsp;</td>
  <th colspan="9" align="center">Date</th>
  <th colspan="9" align="center">Time</th>
  <th colspan="9" align="center">Avg.<br>Weight</th>
  <th colspan="10" align="center">Defects&nbsp;*<br><font size=-2>(Dents/Cracks/Tucks)</font></th>
  <th colspan="9" align="center">Initials</th>
 </tr>';
    foreach my $count ( 1 .. 12 ) {
        push @html, '
 <tr style="height:46px">
  <td colspan="9" align="center">__________</td>
  <td colspan="9" align="center">____:____<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="9" align="center">________&nbsp;mg&nbsp;</td>
  <td colspan="10" align="center">Yes&nbsp;/&nbsp;No</td>
  <td colspan="9" align="center">_________</td>
  <td colspan="8" align="center">&nbsp;</td>
  <td colspan="9" align="center">__________</td>
  <td colspan="9" align="center">____:____<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="9" align="center">________&nbsp;mg&nbsp;</td>
  <td colspan="10" align="center">Yes&nbsp;/&nbsp;No</td>
  <td colspan="9" align="center">_________</td>
 </tr>';
    }
    push @html, $single_blank;
#41 lines
    push @html, '
 <tr style="height:46px">
   <td colspan="22" class="heavy">Average&nbsp;of&nbsp;Averages (this page)</td>
   <td colspan="13" align="center">_________&nbsp;mg</td>
   <td colspan="14" align="center">&nbsp;</td>
   <td colspan="9" class="heavy" align="center">Drums:</td>
   <td colspan="42" align="center">1&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;2&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;3&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;4&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;5&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;6&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;7&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;8&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;9&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;10&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;11&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;12</td>
  </tr>';
    push @html, $single_blank;
    push @html, '
  <tr>
   <td colspan="22"><b>Average&nbsp;Tablet Weight&nbsp;Range</b></td>
   <td colspan="13" class="heavy" align="center" style="vertical-align:bottom">' . $lo_weight . '&nbsp;mg</td>
   <td colspan="10" align="center">and</td>
   <td colspan="13" class="heavy" align="center" style="vertical-align:bottom">' . $hi_weight . '&nbsp;mg</td>
   <td colspan="11" align="right">(&plusmn;5&#37; of Theoretical)<br>Conforms?</td>
   <td colspan="11" align="center">Yes</td>
   <td colspan="11" align="center">No</td>
   <td colspan="9">&nbsp;</td>
  </tr>';
    push @html, $single_blank; 
    push @html, '
 <tr>
  <td colspan="91"><b>&nbsp;&nbsp;*</b> If Capsule Defects are present, contact Quality Control for disposition.</td>
  <td colspan="9" align="center">&nbsp;</td>
 </tr>';
    push @html, $single_blank; 
    push @html, '
 <tr>
  <td colspan="58">If Average Tablet weight is outside of the range, contact Quality Control to address deviation.</td>
  <td colspan="11" class="heavy">QC Deviation:</td>
  <td colspan="11" align="center">#&nbsp;____________</td>
  <td colspan="11" align="center">N/A</td>
  <td colspan="9" align="right">QA-487</td>
 </tr>';
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="30" class="heavy">QC Approval Only</td>
  <td colspan="28">&nbsp;</td>
  <td colspan="11" align="center" class="heavy">Date:&nbsp;_________</td>
  <td colspan="11" align="center" class="heavy">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center" class="heavy">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">&nbsp;</td>
 </tr>';
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $main_tablef;
#12 lines



# # tab_pg10
    $log->log( " > Printing Tablet Yield Log Pg10...", 1 );
    push @html, $main_tableh;
    push @html, '
  <tr>
   <td colspan="100" class="heavy">Tablet Yield Log</td>
 </tr>';
    push @html, '
 <tr>
  <td colspan="100"><li  style="list-style: disc inside">Once the lot is tableted, record the Production
  Yield in the spaces below.</li></td>
 </tr>';
    push @html, $single_blank;  
#5 lines
    foreach my $count ( 1 .. 7 ) {
        push @html, '
 <tr style="height:46px">
  <td colspan="9">&nbsp;</td>
  <td colspan="12" align="center">Container #' . $count . '</td>
  <td colspan="12" align="center">____________&nbsp;Kg</td>
  <td colspan="5" align="center">&nbsp;</td>
  <td colspan="12" align="center">Container #' . ($count+7) . '</td>
  <td colspan="12" align="center">____________&nbsp;Kg</td>
  <td colspan="5" align="center">&nbsp;</td>
  <td colspan="12" align="center">Container #' . ($count+14) . '</td>
  <td colspan="12" align="center">____________&nbsp;Kg</td>
  <td colspan="9">&nbsp;</td>
 </tr>';
    }
    push @html, $single_blank;  
#    push @html, '
# <tr class="double">
#  <td colspan="1">&nbsp;</td>
#  <td colspan="1" class="heavy" align="right">Example:</td>  
#  <td colspan="2" align="center">Total Weight</td>
#  <td colspan="2" align="center"><u><i>&nbsp;&nbsp;' . $runsize . '&nbsp;&nbsp;</i></u>&nbsp;Kg</td>
#  <td colspan="1" align="center">&divide;</td>
#  <td colspan="2" align="center">&nbsp;Avg.&nbsp;Tablet&nbsp;Weight </td>
#  <td colspan="2" align="center"><u><i>&nbsp;&nbsp;&nbsp;' . $fillweight . '&nbsp;&nbsp;&nbsp;</i></u>&nbsp;mg</td>
#  <td colspan="2" align="center">X 1,000,000</td>   
#  <td colspan="2" align="center"><b>= Total Tablets</b></td>
#  <td colspan="3" align="center"><u><i>&nbsp;&nbsp;&nbsp;' . sprintf("%0.f",$runsize/$fillweight*1000000) . '&nbsp;&nbsp;&nbsp;</i></u></td>
# </tr>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="9">&nbsp;</td>
  <td colspan="12" align="center">Total Weight</td>
  <td colspan="12" align="center">_________&nbsp;Kg</td>
  <td colspan="10" align="center">&divide;</td>
  <td colspan="12" align="center">&nbsp;Avg.&nbsp;Tablet&nbsp;Weight </td>
  <td colspan="12" align="center">_________&nbsp;mg</td>
  <td colspan="12" align="center">X 1,000,000</td>  
  <td colspan="12" align="center"><b>= Total Tablets</b></td>
  <td colspan="9" align="center">_________</td>
 </tr>';
    push @html, $single_blank;  
    push @html, '
 <tr>
  <td colspan="9" class="heavy" align="right">Example:</td>  
  <td colspan="12" align="center">Total Weight</td>
  <td colspan="12" align="center"><u><i>&nbsp;&nbsp;&nbsp;&nbsp;500.25&nbsp;&nbsp;&nbsp;&nbsp;</i></u>&nbsp;Kg</td>
  <td colspan="10" align="center">&divide;</td>
  <td colspan="12" align="center">&nbsp;Avg.&nbsp;Tablet&nbsp;Weight </td>
  <td colspan="12" align="center"><u><i>&nbsp;&nbsp;&nbsp;825&nbsp;&nbsp;&nbsp;&nbsp;</i></u>&nbsp;mg</td>
  <td colspan="12" align="center">X 1,000,000</td>   
  <td colspan="12" align="center"><b>= Total Tablets</b></td>
  <td colspan="9" align="center"><u><i>&nbsp;&nbsp;606,364&nbsp;&nbsp;</i></u></td>
 </tr>';
    push @html, $single_blank;
    push @html, '
  <tr>
   <td colspan="21" class="heavy">Theorectical Yield:</td>
   <td colspan="12" align="center" class="heavy">' . $theo_yield . ' tablets</td>
   <td colspan="67">&nbsp;</td>
 </tr>';
 #28 lines
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="21" class="heavy">Theoretical Yield<br>must be between</td>
  <td colspan="12" align="center" class="heavy" style="vertical-align:bottom">' . $lo_yield . ' tablets</td>
  <td colspan="10" align="center">and</td>
  <td colspan="12" align="center" class="heavy" style="vertical-align:bottom">' . $hi_yield . ' tablets</td>
  <td colspan="12" align="right">(&plusmn;3&#37;)<br>Conforms?</td>
  <td colspan="12" align="center">Yes</td>
  <td colspan="12" align="center">No&nbsp;<i>(requires&nbsp;reconciliation)</i></td>
  <td colspan="9">&nbsp;</td>
 </tr>';
    push @html, $single_blank;
    push @html, '
  <tr>
   <td colspan="100">If the Yield is outside of the range, contact Production Manager and Quality for Yield Reconciliation</td>
 </tr>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="30">Tableting started by:</td>
  <td colspan="28">&nbsp;</td>
  <td colspan="11" align="center">Date: _________</td>
  <td colspan="11" align="center">Time: ________<span style="font-weight: 350; font-size: 10px;"> am/pm</span></td>
  <td colspan="11" align="center">Initials: _________</td>
  <td colspan="9">&nbsp;</td>
 </tr>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="30">Appearance/Color OK?</td>
  <td colspan="5" align="center">Yes</td>
  <td colspan="5" align="center">No</td>
  <td colspan="18">&nbsp;</td>  
  <td colspan="11" align="center">Date: _________</td>
  <td colspan="11" align="center">Time: ________<span style="font-weight: 350; font-size: 10px;"> am/pm</span></td>
  <td colspan="11" align="center">Initials: _________</td>
  <td colspan="9">&nbsp;</td>
 </tr>';
    push @html, $single_blank;
    push @html, '
  <tr>
   <td colspan="30">Tablet Coated?</td>
   <td colspan="5" align="center">Yes</td>
   <td colspan="5" align="center">N/A</td>
   <td colspan="18">&nbsp;</td>
   <td colspan="11" align="center">Date:&nbsp;_________</td>
   <td colspan="11" align="center">Time:&nbsp;________<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
   <td colspan="11" align="center">Initials:&nbsp;_________</td>
   <td colspan="9">&nbsp;</td>
  </tr>';
    push @html, $single_blank;
    push @html, '
  <tr>
   <td colspan="30">Hand-sort needed?</td>
   <td colspan="5" align="center">Yes</td>
   <td colspan="5" align="center">No</td>
   <td colspan="18">&nbsp;</td>
   <td colspan="11" align="center">Date:&nbsp;_________</td>
   <td colspan="11" align="center">Time:&nbsp;________<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
   <td colspan="11" align="center">Initials:&nbsp;_________</td>
   <td colspan="9">&nbsp;</td>
 </tr>';
    push @html, $single_blank;
    push @html, '
  <tr>
   <td colspan="30">Hand-sort Loss</td>
   <td colspan="10" align="center">_________&nbsp;Kg</td>
   <td colspan="18">&nbsp;</td>
   <td colspan="11" align="center">Date:&nbsp;_________</td>
   <td colspan="11" align="center">Time:&nbsp;________<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
   <td colspan="11" align="center">Initials:&nbsp;_________</td>
   <td colspan="9">&nbsp;</td>
 </tr>';
    push @html, $single_blank; 
    push @html, '
 <tr>
  <td colspan="30">Other Losses/Leftover Powder</td>
  <td colspan="10" align="center">_________&nbsp;Kg</td>
  <td colspan="18">&nbsp;</td>
  <td colspan="11" align="center">Date:&nbsp;_________</td>
  <td colspan="11" align="center">Time:&nbsp;________<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center">Initials:&nbsp;_________</td> 
  <td colspan="9">&nbsp;</td>
 </tr>';
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, '
  <tr>
   <td colspan="30" class="heavy">QC Batch Yield Reconciliation?</td>
   <td colspan="5" align="center" class="heavy">Yes</td>
   <td colspan="5" align="center" class="heavy">N/A</td>
   <td colspan="18">&nbsp;</td>
   <td colspan="11" align="center" class="heavy">Date:&nbsp;_________</td>
   <td colspan="11" align="center" class="heavy">Time:&nbsp;________<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
   <td colspan="11" align="center" class="heavy">Initials:&nbsp;_________</td>
   <td colspan="9">&nbsp;</td>
  </tr>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="30" class="heavy">QC Approval Only</td>
  <td colspan="28">&nbsp;</td>
  <td colspan="11" align="center" class="heavy">Date:&nbsp;_________</td>
  <td colspan="11" align="center" class="heavy">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center" class="heavy">Initials:&nbsp;_________</td>
  <td colspan="9">&nbsp;</td>
 </tr>';
    push @html, $single_blank;
    push @html, $main_tablef;
#23 lines



# # tab_pg11
    $log->log( " > Printing Post-Tableting pg11...", 1 );
    push @html, $main_tableh;
    push @html, '
 <tr>
   <td colspan="100" class="heavy">Polishing and Dedusting Check</td>
 </tr>';
    push @html, $single_blank; 
    push @html, '
 <tr>
  <td colspan="30">Room #_______ Cleaned</td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="23">&nbsp;</td>
  <td colspan="11" align="center">Date:&nbsp;_________</td>
  <td colspan="11" align="center">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">MF-285</td>
 </tr>';
    push @html, $single_blank; 
    push @html, '
 <tr>
  <td colspan="30">Organic Rinse?</td>
  <td colspan="5" align="center">Yes</td>
  <td colspan="5" align="center">N/A</td>
  <td colspan="51" align="right">&nbsp;</td>
  <td colspan="9" align="right">MF-274</td>
 </tr>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="30">Product Dedusted</td>
  <td colspan="5" align="center">'. $check_box .'</td>
  <td colspan="23">&nbsp;</td>
  <td colspan="11" align="center">Date:&nbsp;_________</td>
  <td colspan="11" align="center">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">&nbsp;</td>
 </tr>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="30">Waste Caps/Powder _______ Kg</td>
  <td colspan="28">&nbsp;</td>
  <td colspan="11" align="center">Date:&nbsp;_________</td>
  <td colspan="11" align="center">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center">Initials:&nbsp;_________</td>
  <td colspan="9">&nbsp;</td>
 </tr>';
    push @html, $single_blank; 
#11 lines
    push @html, '<tbody class="QA">' . $single_blank;
    push @html, '
 <tr>
  <td colspan="30" class="heavy">QC Approval Only</td>
  <td colspan="28">&nbsp;</td>
  <td colspan="11" align="center" class="heavy">Date:&nbsp;_________</td>
  <td colspan="11" align="center" class="heavy">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center" class="heavy">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">&nbsp;</td>
 </tr>';
    push @html, $single_blank . '</tbody>'; 
    push @html, $single_blank;   
    push @html, $single_blank; 
    push @html, $single_blank; 
    push @html, $single_blank; 
    push @html, $single_blank; 
    push @html, $single_blank; 
    push @html, $single_blank; 
    push @html, $single_blank; 
    push @html, $single_blank; 
    push @html, $single_blank; 
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank; 
#16 lines

    push @html, $single_blank;    
    push @html, '
 <tr>
  <td colspan="100" align="left" class="heavy">Tableting Portable Equipment Used</td>
 </tr>';
    push @html, $single_blank; 
    push @html, '
 <tr>
  <td colspan="30">Equipment Used (circle one):</td>
  <td colspan="15" align="center" class="heavy">Colton Mill</td>
  <td colspan="15" align="center" class="heavy">Hammer Mill</td>
  <td colspan="15" align="center" class="heavy">Tablet Press</td>  
  <td colspan="15" align="center" class="heavy">Coating Pan</td> 
  <td colspan="10">&nbsp;</td>
 </tr>';
    push @html, $single_blank; 
    push @html, '
 <tr>
  <td colspan="30">Room #________ Cleaned</td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="23">&nbsp;</td>
  <td colspan="11" align="center">Date:&nbsp;_________</td>
  <td colspan="11" align="center">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">MF-275</td>
 </tr>';
    push @html, $single_blank; 
    push @html, '
 <tr>
  <td colspan="30">Equipment Cleaned</td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="23">&nbsp;</td>
  <td colspan="11" align="center">Date:&nbsp;_________</td>
  <td colspan="11" align="center">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">MF-274</td>
 </tr>';  
    push @html, $single_blank; 
    push @html, '
 <tr>
  <td colspan="30">Organic Rinse?</td>
  <td colspan="5" align="center">Yes</td>
  <td colspan="5" align="center">N/A</td>
  <td colspan="51" align="right">&nbsp;</td>
  <td colspan="9" align="right">MF-274</td>
</tr>';
    push @html, $single_blank; 
    push @html, '
 <tr>
  <td colspan="30">Incoming Weight</td>
  <td colspan="5" align="center">_________&nbsp;Kg</td>
  <td colspan="23">&nbsp;</td>
  <td colspan="11" align="center">Date:&nbsp;_________</td>
  <td colspan="11" align="center">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center">Initials:&nbsp;_________</td>
  <td colspan="9" align="right"></td>
 </tr>';  
    push @html, $single_blank; 
    push @html, '
 <tr>
  <td colspan="30">Outgoing Weight</td>
  <td colspan="5" align="center">_________&nbsp;Kg</td>
  <td colspan="23">&nbsp;</td>
  <td colspan="11" align="center">Date:&nbsp;_________</td>
  <td colspan="11" align="center">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center">Initials:&nbsp;_________</td>
  <td colspan="9" align="right"></td>
 </tr>';
 #14 lines
    push @html, $single_blank; 
    push @html, '<tbody class="QA">' . $single_blank;
    push @html, '
 <tr>
  <td colspan="30" class="heavy">QC Approval Only</td>
  <td colspan="28">&nbsp;</td>
  <td colspan="11" align="center" class="heavy">Date:&nbsp;_________</td>
  <td colspan="11" align="center" class="heavy">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center" class="heavy">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">&nbsp;</td>
 </tr>';
    push @html, $single_blank . '</tbody>'; 
    push @html, $single_blank;
    push @html, $single_blank; 
    push @html, $single_blank; 
    push @html, $single_blank; 
    push @html, $single_blank; 
    push @html, $single_blank; 
    push @html, $single_blank;     
    push @html, '</table><div class="footer"></div>';
#    push @html, $main_tablef;
    return ( 1, @html );
}



sub mmr_pow {
#####
    #
    # Page 2 of MMR - Powder
    #
#####
    my @args      = @_;
    my $form_obj  = shift @args;
    my $ppro_obj  = shift @args;
    my $spec_obj  = shift @args;
    my $batch_obj = shift @args;
    my $runsize   = shift @args;
    my @html      = @args;


# # pow_pg6
    $log->log( " > Printing Powder Clean Check Pg6...", 1 );
    push @html, $main_tableh;
    push @html, '
 <tr>
  <td colspan="100" align="left" class="heavy">Powder Area Clean Check</td>
 </tr>';
    push @html, $single_blank;
    push @html, ' 
 <tr>
  <td colspan="30">Equipment #_______ Cleaned</td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="23">&nbsp;</td>
  <td colspan="11" align="center">Date:&nbsp;_________</td>
  <td colspan="11" align="center">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">PK-535</td>
 </tr>';
    push @html, $single_blank; 
    push @html, '
 <tr>
  <td colspan="30">Room #_______ Cleaned</td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="23">&nbsp;</td>
  <td colspan="11" align="center">Date:&nbsp;_________</td>
  <td colspan="11" align="center">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">PK-535</td>
 </tr>';
    push @html, $single_blank; 
    push @html, '
 <tr>
  <td colspan="30">Organic Rinse?</td>
  <td colspan="5" align="center">Yes</td>
  <td colspan="5" align="center">N/A</td>
  <td colspan="51" align="right">&nbsp;</td>
  <td colspan="9" align="right">MF-274</td>
 </tr>';
    push @html, $single_blank; 
    push @html, '
 <tr>
  <td colspan="30">Appearance/Color OK?</td>
  <td colspan="5" align="center">Yes</th>
  <td colspan="5" align="center">No</th>
  <td colspan="65">&nbsp;</td>
 </tr>';
    push @html, $single_blank; 
#17 lines
    push @html, '<tbody class="QA">' . $single_blank;
    push @html, '
 <tr>
  <td colspan="30" class="heavy">QC Approval Only</td>
  <td colspan="28">&nbsp;</td>
  <td colspan="11" align="center" class="heavy">Date:&nbsp;_________</td>
  <td colspan="11" align="center" class="heavy">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="center" class="heavy">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">&nbsp;</td>
 </tr>';
    push @html, $single_blank . '</tbody>'; 
    push @html, $single_blank;  
    push @html, $single_blank;  
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;    
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;      
    push @html, $single_blank; 
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="30" class="heavy">Excipients Added</td>
  <td colspan="70">&nbsp;</td>
 </tr>';
    push @html, $single_blank;  
    push @html, '
 <tr>
  <td colspan="100" class="heavy" style="vertical-align: top"><b>Formulas:</b></td></tr>
 </tr>';  
    push @html, '
 <tr>
  <td colspan="50"><li style="list-style: disc inside"><b>Total Kg Added</b> = mg Per Unit Added X total capsules</li>
  <li style="list-style: disc inside"><b>mg Per Unit Added</b> = Total Kg Added &divide; total capsules</li></td>
  <td colspan="50">&nbsp;</td>
 </tr>';
    push @html, $single_blank;
    push @html, $single_blank;
#12 lines
    push @html, '
 <tr style="height:37px">
  <th colspan="9">Raw Material</td>
  <th colspan="18">Description</td>
  <th colspan="8">Lot #</td>
  <th colspan="10">Kg added</td>
  <th colspan="10">mg per unit</td>
  <th colspan="12">Purpose</td>
  <th colspan="6">Pre QC by</td>
  <th colspan="6">Added by</td>
  <th colspan="6">Blended by</td>
  <th colspan="6">Posted by</td>
  <th colspan="9">&nbsp;</td>
 </tr>';
    push @html, '
 <tr style="height:37px">
  <td colspan="9" align="center">__________</td>
  <td colspan="18" align="center">__________________________</td>
  <td colspan="8" align="center">__________</td>
  <td colspan="10" align="center">___________&nbsp;Kg</td>
  <td colspan="10" align="center">___________&nbsp;mg</td>
  <td colspan="12" align="center">______________________</td>
  <td colspan="6" align="center">________</td>
  <td colspan="6" align="center">________</td>
  <td colspan="6" align="center">________</td>
  <td colspan="6" align="center">________</td>
  <td colspan="9" align="right">QA-487</td> 
 </tr>';
    push @html, '
 <tr style="height:37px">
  <td colspan="9" align="center">__________</td>
  <td colspan="18" align="center">__________________________</td>
  <td colspan="8" align="center">__________</td>
  <td colspan="10" align="center">___________&nbsp;Kg</td>
  <td colspan="10" align="center">___________&nbsp;mg</td>
  <td colspan="12" align="center">______________________</td>
  <td colspan="6" align="center">________</td>
  <td colspan="6" align="center">________</td>
  <td colspan="6" align="center">________</td>
  <td colspan="6" align="center">________</td>
  <td colspan="9" align="right">QA-487</td> 
 </tr>';
    push @html, '
 <tr style="height:37px">
  <td colspan="9" align="center">__________</td>
  <td colspan="18" align="center">__________________________</td>
  <td colspan="8" align="center">__________</td>
  <td colspan="10" align="center">___________&nbsp;Kg</td>
  <td colspan="10" align="center">___________&nbsp;mg</td>
  <td colspan="12" align="center">______________________</td>
  <td colspan="6" align="center">________</td>
  <td colspan="6" align="center">________</td>
  <td colspan="6" align="center">________</td>
  <td colspan="6" align="center">________</td>
  <td colspan="9" align="right">QA-487</td> 
 </tr>';
 #10 lines
    push @html, $single_blank;      
    push @html, $single_blank;   
    push @html, $single_blank;  
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;    
    push @html, $single_blank;      
    push @html, $single_blank;        
    push @html, $single_blank;   
    push @html, $single_blank;  
    push @html, $single_blank;
    push @html, $single_blank;  
    push @html, '</table><div class="footer"></div>';
#    push @html, $main_tablef; 
    return ( 1, @html );
}




sub mmr_pkg {
#####
    #
    # Page 3 of MMR - Packaging  ## NOT CURRENTLY ACTIVE ##
    #
#####
    my @args      = @_;
    my $form_obj  = shift @args;
    my $ppro_obj  = shift @args;
    my $spec_obj  = shift @args;
    my $batch_obj = shift @args;
    my $runsize   = shift @args;
    my @html      = @args;    

# # pkg_pg1
    $log->log( " > Printing Package Kitting Instructions pkg pg1...", 1 );
    push @html, $main_tableh; 
    push @html, '
 <tr>
  <td colspan="100" align="left" class="heavy">Powder Area Clean Check</td>
 </tr>';    
    push @html, '
 <tbody class="list"><tr>
  <th colspan="83" align="left">Package Kitting Instructions</th>
  <th colspan="5">Read?</th>
  <td colspan="12">&nbsp;</td>  
 </tr>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="83"><li>Print name, sign and initial the Signature table on page 2.</li></td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="12">&nbsp;</td>  
 </tr>'; 
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="83"><li>Review all previous Batch production paperwork and confirm that all entries are complete and accurate.</li></td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="12">&nbsp;</td>   
 </tr>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="83"><li>Start with a clean, undamaged pallet for picking.</li></td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="12">&nbsp;</td>   
 </tr>';  
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="83"><li>Pull all packaging components by item number, in order of the oldest approved component first (top of list downwards).</li></td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="12">&nbsp;</td>   
 </tr>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="83"><li>Verify the ID# and LOT# on the packaging component package against the packaging pick list.</li></td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="12">&nbsp;</td>   
 </tr>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="83"><li>Inspect and carefully remove any loose debris, dust or potential contaminates from packaging component storage containers.</li></td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="12">&nbsp;</td>   
 </tr>'; 
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="83"><li>Circle the lot number and then initial and date Pick List, as the packaging components are pulled.</li></td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="12">&nbsp;</td>   
 </tr>'; 
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="83"><li>Stage the packaging components in the designated staging area.</li></td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="12">&nbsp;</td>    
 </tr></tbody>';
    push @html, $single_blank;   
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="58" class="heavy">All instructions on this page have been read and understood.</td>
  <td colspan="11" align="right">Date:&nbsp;_________</td>
  <td colspan="11" align="right">Time:&nbsp;________<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="right">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">&nbsp;</td> 
 </tr>';
    push @html, $single_blank;  
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="58" class="heavy">Packaging Materials picked by:</td>
  <td colspan="11" align="right">Date:&nbsp;_________</td>
  <td colspan="11" align="right">Time:&nbsp;________<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="right">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">&nbsp;</td> 
 </tr>'; 
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="100" align="left" class="heavy">Packaging Equipment Used</td>
 </tr>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="30">Room #________ Cleaned</td>
  <td colspan="5" align="center">Yes</td>
  <td colspan="5" align="center">N/A</td>
  <td colspan="29" align="right">Date:&nbsp;_________</td>
  <td colspan="11" align="right">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="right">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">MF-240</td>
 </tr>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="30">Counter #________ Used</td>
  <td colspan="28">&nbsp;</td>
  <td colspan="11" align="right">Date:&nbsp;_________</td>
  <td colspan="11" align="right">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="right">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">MF-240</td>
 </tr>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="30">Scale #________ Used</td>
  <td colspan="28">&nbsp;</td>
  <td colspan="11" align="right">Date:&nbsp;_________</td>
  <td colspan="11" align="right">Time:&nbsp;_______<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="right">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">QC-426-IL2</td>
 </tr>';
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;  
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank; 
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $main_tablef;
    
    
    
# # pkg_pg3
    $log->log( " > Printing Pre-Packaging Checklist pkg pg3...", 1 );
    push @html, $main_tableh;
    push @html, '
 <tbody class="list"><tr>
  <th colspan="83" align="left">Pre-Packaging Checklist</th>
  <th colspan="5">Read?</th>
  <td colspan="12">&nbsp;</td>  
 </tr>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="83"><li>Print name, sign and initial the Signature table on page 2.</li></td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="12">&nbsp;</td>  
 </tr>'; 
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="83"><li>Review all previous Batch production paperwork and confirm that all entries are complete and accurate.</li></td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="12">&nbsp;</td>   
 </tr>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="83"><li>Clean and sanitize all table surfaces with cleaner and sanitizer per established procedure.</li></td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="12">&nbsp;</td>  
 </tr>';  
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="83"><li>Sweep and clean floor per established procedure.</li></td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="12">&nbsp;</td>  
 </tr>';   
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="83"><li>Clean all equipment per established procedure.</li></td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="12">&nbsp;</td>   
 </tr>'; 
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="83"><li>Confirm all scoops and other related utensils have been cleaned and sanitized per established procedure.</li></td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="12">&nbsp;</td>   
 </tr>';  
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="83"><li>Obtain the correct label and confirm the quantity. Circle the label lot number, initial, and date the Batch Production Record.</li></td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="12">&nbsp;</td>  
 </tr>'; 
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="83"><li>Set up and confirm packaging machinery is correctly set up (counts, temperatures, rates, etc).</li></td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="12">&nbsp;</td> 
 </tr>'; 
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="83"><li>Lot number and date stamp have been correctly annotated on the label, or labeling equipment has been set up correctly, and confirmed to be accurate.</li></td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="12">&nbsp;</td>  
 </tr>'; 
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="83"><li>Check scale(s) to confirm that the scale calibration is valid and that the scale is set upon a level surface and is level.</li></td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="12">&nbsp;</td>   
 </tr></tbody>'; 
    push @html, $single_blank;   
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="58" class="heavy">All instructions on this page have been read and understood.</td>
  <td colspan="11" align="right">Date:&nbsp;_________</td>
  <td colspan="11" align="right">Time:&nbsp;________<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="right">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">&nbsp;</td> 
 </tr>';
    push @html, $single_blank;
    push @html, '<tbody class="QA list">'.$single_blank; 
    push @html, '
 <tr>
  <th colspan="83" align="left">QC/Supervisor Checklist</th>
  <th colspan="5">Read?</th>
  <td colspan="12">&nbsp;</td>  
 </tr>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="83"><li>Finished product drum(s) are properly staged and all finished product drum labels have the same lot number as the lot number on the Batch Production Record.</li></td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="12">&nbsp;</td>  
 </tr>';   
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="83"><li>All finished product drums have a completed Quality green release sticker.</li></td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="12">&nbsp;</td>  
 </tr>';    
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="83"><li>All component types, lot numbers, and quantities are verified upon review of Batch Production Record.</li></td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="12">&nbsp;</td>  
 </tr>';  
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="83"><li>Labels (if applicable) have been confirmed to be the correct type and lot number.</li></td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="12">&nbsp;</td>  
 </tr>';  
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="83"><li>All packaging equipment is correctly set up (counts, temperatures, rates, etc).</li></td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="12">&nbsp;</td> 
 </tr>'; 
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="83"><li>Lot number and date encoding or Inkjet stamp are confirmed to be accurate.</li></td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="12">&nbsp;</td>   
 </tr>';  
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="83"><li>Packaging room has been properly cleaned and staged per established procedure.</li></td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="12">&nbsp;</td>   
 </tr>'; 
    push @html, $single_blank;   
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="58" class="heavy">QC/Supervisor Approval</td>
  <td colspan="11" align="right">Date:&nbsp;_________</td>
  <td colspan="11" align="right">Time:&nbsp;________<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="right">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">&nbsp;</td> 
 </tr>';
    push @html, $single_blank.'</tbody>';   
    push @html, $main_tablef;
    
 
 
 
# # pkg_pg4
    $log->log( " > Printing Packaging Instructions pkg pg4...", 1 );
    push @html, $main_tableh; 
    push @html, '
 <tbody class="list"><tr>
  <th colspan="83" align="left">Packaging Instructions</th>
  <th colspan="5">Read?</th>
  <td colspan="12">&nbsp;</td>  
 </tr>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="83"><li>Line leads and Supervisors print name, sign and initial the Signature table on page 2.</li></td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="12">&nbsp;</td>   
 </tr>'; 
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="83"><li>Review all previous Batch production paperwork and confirm that all entries are complete and accurate.</li></td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="12">&nbsp;</td>   
 </tr>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="83"><li>In-process Packaging and Monitoring report information has been accurately filled in.</li></td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="12">&nbsp;</td>   
 </tr>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="83"><li>In-process Packaging and Monitoring report will be completed every 15 minutes for the first hour, then every 30 minutes thereafter.</li></td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="12">&nbsp;</td>  
 </tr>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="83"><li>All wasted components (caps, bottles, desiccants, labels and scoops) are to be put into basket to be accounted for.</li></td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="12">&nbsp;</td>  
 </tr>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="83"><li>Pull 2 finished product bottles, at a minimum, for finished product analysis, and as a retention sample.</li></td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="12">&nbsp;</td>  
 </tr>'; 
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="83"><li>Submit finished product sample bottles to laboratory for processing.</li></td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="12">&nbsp;</td>  
 </tr>';  
    push @html, $single_blank;   
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="58" class="heavy">All instructions on this page have been read and understood.</td>
  <td colspan="11" align="right">Date:&nbsp;_________</td>
  <td colspan="11" align="right">Time:&nbsp;________<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="right">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">&nbsp;</td> 
 </tr>';
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;  
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank; 
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;  
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank; 
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;  
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank; 
    push @html, $single_blank;
    push @html, $main_tablef;

    

# # pkg_pg5
    $log->log( " > Printing In-process Packaging and Monitoring pkg pg5...", 1 );
    push @html, $main_tableh;
    push @html, '
 <tr>
  <td colspan="100" class="heavy">In-process Packaging and Monitoring</td>
 </tr>';
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="12">Number of Bottles:</td>
  <td colspan="16">___________</td>
  <td colspan="12">Processing Type:</td>
  <td colspan="16" align="center">'.$check_box.' Original packaging</td>
  <td colspan="6" align="center">'.$check_box.' Rework</td>
  <td colspan="24" align="center">'.$check_box.' Other:&nbsp;________________</td>
  <td colspan="14">&nbsp;</td>
 </tr>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="12">Label Item Number:</td>
  <td colspan="16">___________</td>
  <td colspan="12">Label Lot Numbers:</td>
  <td colspan="16">___________________________</td>
  <td colspan="12">Desiccant Size:</td>
  <td colspan="16">_________________</td>
  <td colspan="16">&nbsp;</td>
 </tr>';
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, '
 <tr>
  <th colspan="6">Time</th>
  <th colspan="6">Cap/Tab<br>Count</th>
  <th colspan="5">Correct<br>Label</th>
  <th colspan="5">Correct<br>Bottle</th>
  <th colspan="5">Correct<br>Cap</th>
  <th colspan="6">Inner Seal<br>Check</th>
  <th colspan="6">Outer Seal<br>Check</th>
  <th colspan="6">Label<br>Quality Check</th>
  <th colspan="6">Correct<br>Mfg/Expire<br>Date</th>
  <th colspan="6">Desiccant<br>Added (size)</th>
  <th colspan="6">Cotton/Rayon<br>Added (type)</th>
  <th colspan="5">Correct<br>Scoop</th>
  <th colspan="5">Correct<br>Carton</th>
  <th colspan="12">Verified<br>by</th>
  <th colspan="12">Date</th>
  <th colspan="3">&nbsp;</td>
 </tr>'; 
    for (my $i = 0; $i < 18; $i++) {
        push @html, '
 <tr height="35px">
  <td colspan="6" align="right" class="QA">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="6" class="QA">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
  <td colspan="5" class="QA">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
  <td colspan="5" class="QA">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
  <td colspan="5" class="QA">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
  <td colspan="6" class="QA">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
  <td colspan="6" class="QA">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
  <td colspan="6" class="QA">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
  <td colspan="6" class="QA">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
  <td colspan="6" class="QA">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
  <td colspan="6" class="QA">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
  <td colspan="5" class="QA">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
  <td colspan="5" class="QA">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
  <td colspan="12" class="QA">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
  <td colspan="12" class="QA">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
  <td colspan="3">&nbsp;</td>
 </tr>';         
    }
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $main_tablef;



# # pkg_pg6
    $log->log( " > Printing Label Counts pkg pg6...", 1 );
    push @html, $main_tableh;
    push @html, '
  <tr>
   <td colspan="88" class="heavy">Label Counts</td>
   <td colspan="12" align="right">PK-534</td>
 </tr>';
    push @html, $single_blank;
    push @html, ' 
 <tr>
  <td colspan="30">Number of Labels In:</td>
  <td colspan="28">______________</td>
  <td colspan="11" align="right">Date:&nbsp;_________</td>
  <td colspan="11" align="right">Time:&nbsp;________<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="right">Initials:&nbsp;_________</td>
  <td colspan="9" align="right"></td>
 </tr>';
    push @html, $single_blank;
    push @html, ' 
 <tr>
  <td colspan="30">Labels Used:</td>
  <td colspan="28">______________</td>
  <td colspan="11" align="right">Date:&nbsp;_________</td>
  <td colspan="11" align="right">Time:&nbsp;________<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="right">Initials:&nbsp;_________</td>
  <td colspan="9" align="right"></td>
 </tr>';
    push @html, $single_blank;
    push @html, ' 
 <tr>
  <td colspan="30">Labels Damaged/Wasted:</td>
  <td colspan="28">______________</td>
  <td colspan="11" align="right">Date:&nbsp;_________</td>
  <td colspan="11" align="right">Time:&nbsp;________<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="right">Initials:&nbsp;_________</td>
  <td colspan="9" align="right"></td>
 </tr>';
    push @html, $single_blank;
    push @html, ' 
 <tr>
  <td colspan="30">Labels Out:</td>
  <td colspan="28">______________</td>
  <td colspan="11" align="right">Date:&nbsp;_________</td>
  <td colspan="11" align="right">Time:&nbsp;________<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="right">Initials:&nbsp;_________</td>
  <td colspan="9" align="right"></td>
 </tr>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="30">To Location:</td>
  <td colspan="29">______________</td>
  <td colspan="11" align="right">Date:&nbsp;_________</td>
  <td colspan="11" align="right">Time:&nbsp;________<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="right">Initials:&nbsp;_________</td>
  <td colspan="9" align="right"></td>
 </tr>';
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;  
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank; 
    push @html, $single_blank;
    push @html, $single_blank;  
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank; 
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;  
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank; 
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;  
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank; 
    push @html, $main_tablef;    
    
    
    
# # pkg_pg7
    $log->log( " > Printing Bulk Packaging Area Clean Check pkg pg7...", 1 );
    push @html, $main_tableh;
    push @html, '
 <tr>
  <td colspan="100" align="left" class="heavy">Bulk Packaging Area Clean Check (if applicable)</td>
 </tr>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="30">Bulk Packaging Room Cleaned</td>
  <td colspan="5" align="center">Yes</td>
  <td colspan="5" align="center">N/A</td>
  <td colspan="29" align="right">Date:&nbsp;_________</td>
  <td colspan="11" align="right">Time:&nbsp;________<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="right">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">MF-219</td>
 </tr>';
   push @html, $single_blank;
    push @html, ' 
 <tr>
  <td colspan="30">Scale #_______ Cleaned and Calibrated</td>
  <td colspan="5" align="center">Yes</td>
  <td colspan="5" align="center">N/A</td>
  <td colspan="29" align="right">Date:&nbsp;_________</td>
  <td colspan="11" align="right">Time:&nbsp;________<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="right">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">MF-</td>
 </tr>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="30">Allergen Test needed?</td>
  <td colspan="5" align="center">Yes</td>
  <td colspan="5" align="center">No</td>
  <td colspan="29" align="right">Date:&nbsp;_________</td>
  <td colspan="11" align="right">Time:&nbsp;________<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="right">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">QA-498<br>MF-274</td>
</tr>';
    push @html, $single_blank;
    push @html, '<tbody class="QA">'.$single_blank;
    push @html, '
 <tr>
  <td colspan="30" class="heavy">QC/Supervisor Approval</td>
  <td colspan="28">&nbsp;</td>
  <td colspan="11" align="right" class="heavy">Date:&nbsp;_________</td>
  <td colspan="11" align="right" class="heavy">Time:&nbsp;________<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="right" class="heavy">Initials:&nbsp;_________</td>
  <td colspan="9" align="right"></td>
 </tr>';
    push @html, $single_blank.'</tbody>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="18" class="heavy">Bulk Packaging Instructions</td>
 </tr>';
    push @html, $single_blank;
    push @html, '
 <tbody class="list"><tr>
  <th colspan="83" align="left">Instructions</th>
  <th colspan="5">Read?</th>
  <td colspan="12">&nbsp;</td>  
 </tr>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="83"><li>Print name, sign and initial the Signature table on page 2.</li></td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="12">&nbsp;</td>  
 </tr>'; 
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="83"><li>Bulk Packaging Room clean check has been performed and approved by Quality.</li></td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="12">&nbsp;</td> 
 </tr>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="83"><li>Ensure that all proper Personal Protective Equipment (PPE) is being utilized and worn correctly.</li></td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="12">&nbsp;</td>  
 </tr>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="83"><li>Finished product has been properly identified and staged.</li></td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="12">&nbsp;</td> 
 </tr>'; 
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="83"><li>When bulking capsules or tablets, utilize the Average of Average Weights from the 10 Count Average Weight Logs.</li></td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="12">&nbsp;</td> 
 </tr>'; 
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="83"><li>Pull the proper bulk container, as listed in the Tier 3 Batch Record section.</li></td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="12">&nbsp;</td>  
 </tr>'; 
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="83"><li>Line bulk container with 2 4-mil palstic bags.</li></td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="12">&nbsp;</td>
 </tr>'; 
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="83"><li>After lining, set the lined container on the scale and tare.</li></td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="12">&nbsp;</td>
 </tr>'; 
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="83"><li>When proper amount is weighed, zip tie both the inner and outer bags.</li></td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="12">&nbsp;</td> 
 </tr>'; 
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="83"><li>Label bulk containers.</li></td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="12">&nbsp;</td> 
 </tr>'; 
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="83"><li>Accurately complete the Bulk Packaging Form.</li></td>
  <td colspan="5" align="center">'.$check_box.'</td>
  <td colspan="12">&nbsp;</td> 
 </tr>'; 
    push @html, $single_blank;   
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="58" class="heavy">All instructions on this page have been read and understood.</td>
  <td colspan="11" align="right">Date:&nbsp;_________</td>
  <td colspan="11" align="right">Time:&nbsp;________<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="right">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">&nbsp;</td> 
 </tr>';  
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $main_tablef;
    


# # pkg_pg8
    $log->log( " > Printing Bulk Packaging Form pkg pg8...", 1 );
    push @html, $main_tableh;
    push @html, '
 <tr>
  <td colspan="100" class="heavy">Bulk Packaging Form</td>
 </tr>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="6"&nbsp;</td>
  <th colspan="11">Container</th>
  <th colspan="11">Tare Weight</th>
  <th colspan="11">Net Weight</th>
  <th colspan="11">Average<br>Cap/Tab Weight</th>
  <th colspan="11">Number of<br>Cap/Tab</th>
  <th colspan="11">Date</th>
  <th colspan="11">Time </th>
  <th colspan="11">Initials</th>
  <td colspan="6">&nbsp;</td>
 </tr>';

    foreach my $count ( 1 .. 18 ) {
        push @html, '
 <tr height="35px">
  <td colspan="6">&nbsp;</td> 
  <td colspan="11" class="QA">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
  <td colspan="11" class="QA" align="right">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Kg</td>
  <td colspan="11" class="QA" align="right">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Kg</td>
  <td colspan="11" class="QA" align="right">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;mg</td>
  <td colspan="11" class="QA">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
  <td colspan="11" class="QA">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
  <td colspan="11" align="right" class="QA">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" class="QA">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
  <td colspan="6">&nbsp;</td>
 </tr>';
    }
    push @html, $single_blank;
    push @html, $single_blank;  
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank; 
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;  
    push @html, $main_tablef;
    
    
    
# # pkg_pg9    
    $log->log( " > Printing Pallet Form pkg pg9...", 1 );
    push @html, $main_tableh;
    push @html, '
 <tr>
  <td colspan="100" class="heavy">Pallet Form</td>
 </tr>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="6">&nbsp;</td>
  <td colspan="12" class="heavy">Company:</td>
  <td colspan="12">_______________</td>
  <td colspan="6">&nbsp;</td>
  <td colspan="12" class="heavy">Order Size:</td>
  <td colspan="12">_______________</td>  
  <td colspan="6">&nbsp;</td>
  <td colspan="12" class="heavy">Lot Number:</td>
  <td colspan="12">_______________</td>
  <td colspan="10">&nbsp;</td>
 </tr>';
    push @html, $single_blank;    
    push @html, '
 <tr>
  <td colspan="6">&nbsp;</td>
  <td colspan="12" class="heavy">Pallet Number:</td>
  <td colspan="12">_______________</td>
  <td colspan="36">&nbsp;</td>
  <td colspan="12" class="heavy">Pallet Number:</td>
  <td colspan="12">_______________</td>
  <td colspan="10">&nbsp;</td>   
 </tr>';
    push @html, $single_blank;    
    push @html, '
 <tr>
  <td colspan="6">&nbsp</td>
  <td colspan="12" class="heavy QA">Master Box #</td>
  <td colspan="12" class="heavy QA">Weight</td>
  <td colspan="18" class="heavy QA">Amount of Bottles</td>
  <td colspan="18">&nbsp;</td>
  <td colspan="12" class="heavy QA">Master Box #</td>
  <td colspan="12" class="heavy QA">Weight</td>
  <td colspan="10" class="heavy QA">Amount of Bottles</td>
 </tr>';
    for (my $i = 0; $i < 9; $i++) {
                push @html, '
 <tr style="height: 28px">
  <td colspan="6">&nbsp</td>
  <td colspan="12" class="QA">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
  <td colspan="12" class="QA">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
  <td colspan="18" class="QA">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
  <td colspan="18">&nbsp;</td>
  <td colspan="12" class="QA">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
  <td colspan="12" class="QA">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
  <td colspan="10" class="QA">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
 </tr>';
    }
    push @html, '
 <tr style="height: 28px">
  <td colspan="6" class="heavy QA">Totals</td>
  <td colspan="12" class="QA">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
  <td colspan="12" class="QA">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
  <td colspan="18" class="QA">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
  <td colspan="12">&nbsp;</td>
  <td colspan="6" class="heavy QA">Totals</td>
  <td colspan="12" class="QA">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
  <td colspan="12" class="QA">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
  <td colspan="10" class="QA">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
 </tr>';
    push @html, '<tbody class="QA">'.$single_blank;
    push @html, '
  <tr>
   <td colspan="58">Line Verification</td>
   <td colspan="11" align="right">Date:&nbsp;_________</td>
   <td colspan="11" align="right">Time:&nbsp;________<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
   <td colspan="11" align="right">Initials:&nbsp;_________</td>
   <td colspan="9">&nbsp;</td>
 </tr>'; 
    push @html, $single_blank;
    push @html, '
  <tr>
   <td colspan="58">Supervisor Verification</td>
   <td colspan="11" align="right">Date:&nbsp;_________</td>
   <td colspan="11" align="right">Time:&nbsp;________<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
   <td colspan="11" align="right">Initials:&nbsp;_________</td>
   <td colspan="9">&nbsp;</td>
 </tr>';
    push @html, $single_blank;
    push @html, '
  <tr>
   <td colspan="58">Shipping Verification</td>
   <td colspan="11" align="right">Date:&nbsp;_________</td>
   <td colspan="11" align="right">Time:&nbsp;________<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
   <td colspan="11" align="right">Initials:&nbsp;_________</td>
   <td colspan="9">&nbsp;</td>
 </tr>';
    push @html, $single_blank; 
    push @html, '</tbody><tbody class="QA">'.$single_blank;    
    push @html, '
 <tr>
  <td colspan="12">Total&nbsp;Number<br>of Boxes:</td>
  <th colspan="16">__________</th>
  <td colspan="5">&nbsp;</td>  
  <td colspan="12">Total&nbsp;Amount<br>of Weight:</td>
  <th colspan="16" align="right">__________&nbsp;Kg</th>
  <td colspan="5">&nbsp;</td>  
  <td colspan="12">Total&nbsp;Number<br>of Pallets:</td>
  <th colspan="16">__________</th>
  <td colspan="6">&nbsp;</td>  
 </tr>';
    push @html, $single_blank;    
    push @html, '
 <tr>
  <td colspan="30" class="heavy">Total Number of Bottles produced for Batch:</td>
  <td colspan="28">_______________</td>
  <td colspan="6">&nbsp;</td>
 </tr>';
    push @html, $single_blank.'</tbody>';    
    push @html, '
 <tr style="height: 28px">
  <td colspan="12" class="heavy QA">Date</td>
  <td colspan="12" class="heavy QA">Run #</td>
  <td colspan="12" class="heavy QA">Bottles Produced</td>
  <td colspan="16">&nbsp</td>
  <td colspan="12" class="heavy QA">Date</td>
  <td colspan="12" class="heavy QA">Bottles Labeled</td>
  <td colspan="12" class="heavy QA">Bottles Unlabeled</td>
  <td colspan="12" class="heavy QA">Bottles Shipped</td>
 </tr>';
    for (my $i = 0; $i < 5; $i++) {
                push @html, '
 <tr style="height: 28px">
  <td colspan="12" class="QA">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
  <td colspan="12" class="QA">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
  <td colspan="12" class="QA">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
  <td colspan="16">&nbsp</td>
  <td colspan="12" class="QA">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
  <td colspan="12" class="QA">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
  <td colspan="12" class="QA">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
  <td colspan="12" class="QA">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
 </tr>';
    }
    push @html, '<tbody class="QA">';    
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="12">Pallet Configuration:</td>
  <td colspan="11">__________</td>
  <td colspan="11">&nbsp;</td>
  <td colspan="12">Pallet Dimensions:</td>
  <td colspan="11">__________</td>
  <td colspan="10">&nbsp;</td>
  <td colspan="12">Total Weight<br>plus 45 lbs:</td>
  <td colspan="11">__________</td>
  <td colspan="10">&nbsp;</td>  
 </tr>';
    push @html, $single_blank.'</tbody>'; 
    push @html, $main_tablef;



# # pkg_pg10
    $log->log( " > Printing Shipping Checklist pkg pg10...", 1 );
    push @html, $main_tableh;
    push @html, $single_blank;    
    push @html, '
 <tr>
  <td colspan="100" class="heavy">Shipping Checklist</td>
 </tr>';
    push @html, $single_blank;
    push @html, '
<tr>
  <td colspan="58"><li>Verify the quantity shipping from the packing slip.</li></td>
  <td colspan="11" align="right">Date:&nbsp;_________</td>
  <td colspan="11" align="right">Time:&nbsp;________<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="right">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">&nbsp;</td>  
 </tr>'; 
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="58"><li>Gather all Pallet Sheets and attach to Batch Record.</li></td>
  <td colspan="11" align="right">Date:&nbsp;_________</td>
  <td colspan="11" align="right">Time:&nbsp;________<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="right">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">&nbsp;</td>  
 </tr>';  
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="58"><li>Fill out Date, Number of bottles shipped and initial the Packaging Master.</li></td>
  <td colspan="11" align="right">Date:&nbsp;_________</td>
  <td colspan="11" align="right">Time:&nbsp;________<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="right">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">&nbsp;</td>  
 </tr>';
    push @html, $single_blank;
    push @html, '
 <tr>
  <td colspan="58"><li>Put all paperwork in top bin on Shipping Desk for finalization.</li></td>
  <td colspan="11" align="right">Date:&nbsp;_________</td>
  <td colspan="11" align="right">Time:&nbsp;________<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
  <td colspan="11" align="right">Initials:&nbsp;_________</td>
  <td colspan="9" align="right">&nbsp;</td>  
 </tr></tbody>';     
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;  
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank; 
    push @html, $single_blank;
    push @html, $single_blank;  
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank; 
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;  
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank; 
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;  
    push @html, $single_blank;
    push @html, $single_blank;
    push @html, $single_blank;    
    push @html, '</table><div class="footer"></div>';  
    return ( 1, @html );  
}




sub make_cofm {
#####
    #
    # make_cofm - Make Certificate of Manufacture for Product ID
    #
#####
    local $| = 1;
    my @args      = @_;
    my $product   = $args[0];
    my $form_form = FormFormula->new($product);
    my $ppro_form = PProFormula->new($product);
    my ( $batchnumber, $index );

    $log->log( ">>>make_cofm $product", 2 );

    while (1) {
        my (%list);
        print "\nBR and LOT number(s) in ProcessPro (only the past year is displayed):\n";
        print "\n        BR\tLOT NO  \tMFG DATE\n";
        foreach my $count ( 0 .. $#{ $ppro_form->{'Wono'} } ) {
            my @dt       = split( q{ }, $ppro_form->{'Wono'}[$count][3] );
            my @end_date = split( q{-}, $dt[0] );
            my $mfg_date;
            if (   ( $end_date[0] < 2000 )
                && ( $ppro_form->{'Wono'}[$count][4] == 5 ) ) {
                $mfg_date = "- Voided";
            } elsif ( $ppro_form->{'Wono'}[$count][4] == 4 ) {
                $mfg_date = sprintf "%02d/%02d/%04d", $end_date[1], $end_date[2], $end_date[0];
            } else {
                $mfg_date = "- Open";
            }
            $list{ int( $ppro_form->{'Wono'}[$count][1] ) } = $count;
            if (( $end_date[0] > ($date[5] + 1898) )||( $end_date[0] < 2000 )) {
                printf "%s\t%s\t%s\n", $ppro_form->{'Wono'}[$count][1], $ppro_form->{'Wono'}[$count][2], $mfg_date;
            }
        }

        $batchnumber = prompt("\tEnter Batch Record to use <x - cancel>: ");
        if ($batchnumber =~ m/^x$/i) {
            $log->log(">> Document canceled!",1);
            return 0;
        } elsif ( !exists $list{$batchnumber} ) {
            print "\n\t$batchnumber is invalid!\n";
            next;
        }
        $index = $list{$batchnumber};
        last;
    }

    if ( $ppro_form->{'Wono'}[$index][4] < 4 ) {
        $log->append("WARNING $bell: This BR has not been closed. Changes may still occur!", 1 );
    }

    my $batch = Batch->new($batchnumber);
    $log->log( " > BR used:\t$batch->{'BatchRecord'}",         1 );

    write_cofm( $form_form, $ppro_form, $batch );

    return 1;
}

sub write_cofm {
#####
    #
    # Write_cofm - Output HTML to file
    #
#####
    local $| = 1;
    my @args      = @_;
    my $form_obj  = $args[0];
    my $ppro_obj  = $args[1];
    my $batch_obj = $args[2];

    my $path = "$ENV{USERPROFILE}\\Desktop\\";
    $ppro_obj->{'FormulaCode'} =~ s/\./ /g;
    my $file = sprintf "%d %s %s%s", $batch_obj->{'Lot'}, $ppro_obj->{'FormulaCode'}, $ppro_obj->{'Description'}, " CofM.xlsx";
    $file =~ s|\#||g;
    $file =~ s|/|_|g;
    $file =~ s|^a-zA-Z\d\s[()+,-]|_|g;

    my $xcel_obj = Excel::Writer::XLSX->new($path.$file);
    if (!$xcel_obj) {
        $log->log(" ! Unable to create $path"."$file!\n$!\n", 1);
        return 0;
    }
    my ( $complete ) = cofm_excel( $form_obj, $ppro_obj, $batch_obj, $xcel_obj );
    if ( !$complete ) {
        $log->log( " !!Document failed to generate, due to errors ##", 1 );
        return 0;
    }
    $xcel_obj->close;
    $log->log( ">> $path"."$file completed!", 1 );
    return 1;
}

sub cofm_excel {
#####
    #
    # Page 1 of CofM - Cover Sheet and Signoffs
    #
#####
    local $| = 1;  
    use Excel::Writer::XLSX;
  
    my @args      = @_;
    my $form_obj  = $args[0];
    my $ppro_obj  = $args[1];
    my $batch_obj = $args[2];
    my $xcel      = $args[3];

    $xcel->set_tempdir("$ENV{USERPROFILE}\\Desktop\\");
    my $wrksheet = $xcel->add_worksheet();
    $wrksheet->set_landscape();
    $wrksheet->set_paper(1);
    $wrksheet->set_margins(.25);    
    $wrksheet->fit_to_pages( 1,0);
        
    my $heading = $xcel->add_format(
        font=>'Book Antiqua',
        size=>12,
        bold=>0 );
    my $headingb = $xcel->add_format(
        font=>'Book Antiqua',
        size=>12,
        bold=>1 );
    my $headingt = $xcel->add_format(
        font=>'Book Antiqua',
        size=>24,
        bold=>1 );
    my $parag = $xcel->add_format(
        font=>'Courier New',
        size=>8,
        text_wrap=>1 );
    my $body = $xcel->add_format(
        font=>'Courier New',
        size=>8 );
    my $bodyr = $xcel->add_format(
        font=>'Courier New',
        size=>8,
        align=>'right' );
    my $bodyc = $xcel->add_format(
        font=>'Courier New',
        size=>8,
        align=>'center' );
    my $bodycb = $xcel->add_format(
        font=>'Courier New',
        size=>8,
        align=>'center',
        bold=>1 );
    my $bodyb = $xcel->add_format(
        font=>'Courier New',
        size=>8,
        bold=>1 );
    my $bodybr = $xcel->add_format(
        font=>'Courier New',
        size=>8,
        bold=>1,
        align=>'right');
    my $sigline = $xcel->add_format(
        font=>'Courier New',
        size=>12,
        bold=>1 );
    my $bborder = $xcel->add_format(
        bottom=>1,
        bottom_color=>'Black' );

    my ( %allergens, @html, $complete, %prodtotal, %alloctotal, @list, %prawindex, %frawindex, %allraws, %objindex);

    if (   ( $form_obj->{'ServingDesc'} eq q{} )
        || ( $form_obj->{'ServingSize'} eq q{} ) ) {
        $log->log( '!! There is no Cap/Tab information in Formulator!', 1 );
        return 0;
    }
    my $dose = sprintf "%s %s %s", $form_obj->{'ServingSize'}, $form_obj->{'Appearance'}, $form_obj->{'ServingDesc'};

    foreach my $count ( 0 .. $#{ $form_obj->{'RawMaterials'} } ) {
        $frawindex{ $form_obj->{'RawMaterials'}[$count][2] }   = $count;
    }
#FormulaTool->vardump(%allraws);

    foreach my $count ( 0 .. $#{ $ppro_obj->{'Details'} } ) {
        $prawindex{ FormulaTool->despace($ppro_obj->{'Details'}[$count][0]) }   = $count;
    }
#FormulaTool->vardump(%allraws);

    foreach my $count ( 0 .. $#{ $batch_obj->{'BatchDetails'} } ) {
        if ( $batch_obj->{'BatchDetails'}[$count][0] eq 'ZZ' ) {
            $alloctotal{ FormulaTool->despace( $batch_obj->{'BatchDetails'}[$count][4] ) } += $batch_obj->{'BatchDetails'}[$count][6];
        } elsif ( $batch_obj->{'BatchDetails'}[$count][0] eq 'MI' ) {
            $prodtotal{ FormulaTool->despace( $batch_obj->{'BatchDetails'}[$count][4] ) } += $batch_obj->{'BatchDetails'}[$count][5];
        } else { next; }
        if (! exists $prawindex{ FormulaTool->despace($batch_obj->{'BatchDetails'}[$count][4] ) } ) {
            $prawindex{ FormulaTool->despace($batch_obj->{'BatchDetails'}[$count][4] ) } = $count;
        }
    }
#FormulaTool->vardump(%allraws); 

    foreach my $count ( 0 .. $#{ $form_obj->{'Objectives'} } ) {
        if ( defined($form_obj->{'Objectives'}[$count][4]) ) {
            $objindex{$form_obj->{'Objectives'}[$count][4]} = $count;
        }
    }
#FormulaTool->vardump(%objindex); #return;

    my @dt       = split( q{ }, $batch_obj->{'CloseDate'} );
    my @end_date = split( q{-}, $dt[0] );
    my $close_date = sprintf "%02d/%02d/%04d", $end_date[1], $end_date[2], $end_date[0];

    my $rawcount = 0;
    my $atotal   = 0;
    my $ptotal   = 0;
    my $ftotal   = 0;

    foreach my $index ( keys %frawindex ) {
        my ( $inert, $objdesc, $target, $tunit, $activity );
        $index =~ s/\s\s+//g;
        $list[$rawcount][0] = $index;
        my $praw_obj = PProRaw->new($index);
        $praw_obj->{'Description'} =~ s/(.+)\^in.+/$1/i;
        $list[$rawcount][1] = $praw_obj->{'Description'};
        my $fraw_obj = FormRaw->new($index);
        if (exists $objindex{$index}) {
            $log->log(" > Checking Objectives for $index in Formulator...",1);
            $objdesc  = $form_obj->{'Objectives'}[$objindex{$index}][0];
            $target   = $form_obj->{'Objectives'}[$objindex{$index}][1];
            $tunit    = $form_obj->{'Objectives'}[$objindex{$index}][2];
            foreach my $nucount ( 0 .. $#{ $fraw_obj->{'Nutrients'} } ) {
                $log->log(" > \t\t '$fraw_obj->{'Nutrients'}[$nucount][2]' in Formulator...",1);
                if ($objdesc eq $fraw_obj->{'Nutrients'}[$nucount][2]) {
                    $activity = sprintf "%.4f", ( $fraw_obj->{'Nutrients'}[$nucount][1] / $fraw_obj->{'ServingSize'} );
                    last;
                }
            }
        } else {
                $log->log(" > No Objectives in Formulator for $index...",1);
                $target   = 'other';
                $tunit    = q{};
                $activity = 0;
                $objdesc  = q{};
        }
        my $increment = ( length $objdesc < 15 ) ? length $objdesc : 15;
        $objdesc =~ s/CS\s\w+\s//;
        $list[$rawcount][2] = ( substr( lc $objdesc, 0, $increment ) eq substr( lc $list[$rawcount][1], 0, $increment ) ? q{} : $objdesc );
        $list[$rawcount][3] = ( int $target > 0 ? ( $target - int $target > 0 ? sprintf "%0.2f", $target : sprintf "%0.0f", $target ) : $target );
        $list[$rawcount][4] = $tunit;
        $list[$rawcount][5] = sprintf "%0.2f", $form_obj->{'RawMaterials'}[ $frawindex{$index} ][3];
        $ftotal += $list[$rawcount][5];
        $list[$rawcount][6] = $form_obj->{'RawMaterials'}[ $frawindex{$index} ][5];
        $list[$rawcount][7] = ($form_obj->{'RawMaterials'}[ $frawindex{$index} ][8] <=0)?q{}:sprintf "%0.1f", $form_obj->{'RawMaterials'}[ $frawindex{$index} ][8];# q{};# $ovg;

        $rawcount++;
    }
    foreach my $index ( keys %prawindex ) {
        if (! exists $frawindex{$index}) {
            my $fraw_obj = FormRaw->new($index);
            my $praw_obj = PProRaw->new($index);
            $log->log(" > Checking for '$index $praw_obj->{'Description'}' in Formulator...",1);
            if (undef == $fraw_obj->{'Description'}) {
                $praw_obj->{'Description'} =~ s/(.+)\^in.+/$1/i;
                $fraw_obj->{'Description'} = "** " . $praw_obj->{'Description'};
            } else {
                $fraw_obj->{'Description'} =~ s/(.+)\^in.+/$1/i;
            }
            $list[$rawcount][1] = $fraw_obj->{'Description'};
            $list[$rawcount][0] = $index;
            $list[$rawcount][5] = "Not in Frm";
            $list[$rawcount][8] = ( $prodtotal{$index} == 0 ? $prodtotal{$index} : sprintf "%0.2f", $prodtotal{$index} );
            $list[$rawcount][9] = 'Kg';
            $rawcount++;
        }
    }

    foreach my $index ( 0 .. $#list ) {
        $log->log(" > Checking $list[$index][0] for Production amounts...",1);
        if ((! exists $prodtotal{$list[$index][0]}) && (! exists $alloctotal{$list[$index][0]})) {
            #$list[$index][1] = "** ".$list[$index][1];
            $list[$index][8] = 'Not on BR';
        } elsif ((! exists $prodtotal{$list[$index][0]}) && (exists $alloctotal{$list[$index][0]})) {
            $list[$index][8] = 'Incomplete';
            $list[$index][9] = sprintf "S/B %0.2f",$alloctotal{$list[$index][0]};
            $atotal += $alloctotal{$list[$index][0]};
        } else {
            $list[$index][8] = ( $prodtotal{$list[$index][0]} == 0 ? $prodtotal{$list[$index][0]} : sprintf "%0.2f", $prodtotal{$list[$index][0]} );
            $list[$index][9] = 'Kg';
            $ptotal += $prodtotal{$list[$index][0]};
            foreach my $count (0 .. $#{ $ppro_obj->{"Details"} }) {
                if ($list[$index][0] == FormulaTool->despace($ppro_obj->{"Details"}[$count][0]) ) {
                    my $check1 = $alloctotal{$list[$index][0]};
                    my $check2 =  $prodtotal{$list[$index][0]};
                    if ($check2 != 0) {
                        if (($check1/$check2) > 1.015 || ($check1/$check2) < 0.985) {
                            $list[$index][9] = sprintf "S/B %.2f",$check1;
                        }
                    } else {
                            $list[$index][8] = 'Not on BR';
                    }
                }
            }
        }
        $list[$index][0] = (exists $ppro_obj->{'Allergens'}{ $list[$index][0] })? $list[$index][0]." *" : $list[$index][0];
    }

    if ($ptotal == 0) {
        $ptotal = sprintf "S/B %0.2f", $atotal;
    } else {
        $ptotal = sprintf "%0.2f", $ptotal;
    }
    $ftotal = sprintf "%0.2f", $ftotal;
    
    my $title = $batch_obj->{'Lot'} . q{ } . $batch_obj->{'FormulaCode'} . q{ } . $ppro_obj->{'Description'};
    $title =~ s/\./ /gm;

    $log->log( " > Printing CofM...", 1 );
    $wrksheet->set_header( '&R&07Page &P of &N', 0.05, { scale_with_doc => 0, align_with_margins => 0 } );
    $wrksheet->set_column( 'A:A',10.7);
    $wrksheet->set_column( 'B:D',8.5);
    $wrksheet->set_column( 'E:E',32.5);
    $wrksheet->set_column( 'F:F',21.7);
    $wrksheet->set_column( 'G:G',14);
    $wrksheet->set_column( 'H:H',3.7);
    $wrksheet->set_column( 'I:I',10.7);
    $wrksheet->set_column( 'J:J',3.7);
    $wrksheet->set_column( 'K:K',10.7);
    $wrksheet->set_column( 'L:L',3.7);
    $wrksheet->set_column( 'M:M',10.7);
    $wrksheet->set_column( 'N:N',3.7);
    $wrksheet->set_row( 0, 55);
    $wrksheet->insert_image('A1', 'cns_logo.jpg', 3, 5, 3, 3); 
#    $wrksheet->merge_range_type( 'string', 'A1:J1', 'Columbia Nutritional, LLC.', $heading);
    $wrksheet->merge_range_type( 'string', 'A2:J2', '6317 NE 131st Ave, Suite 103', $heading);
    $wrksheet->merge_range_type( 'string', 'A3:J3', 'Vancouver, WA  98682', $heading);
    $wrksheet->merge_range_type( 'string', 'A4:J4', 'ph. (360) 737-9966 fax (360) 737-9772', $heading);
    $wrksheet->write_blank( 4, 0, $heading);
    $wrksheet->set_row( 5, 30);
    $wrksheet->merge_range_type( 'string', 'A6:J6', 'Certificate of Manufacture', $headingt );
    $wrksheet->write_blank( 6, 0, $heading);
    $wrksheet->merge_range_type( 'string', 'A8:B8', 'Master Number:', $heading); $wrksheet->merge_range_type( 'string', 'C8:E8', $batch_obj->{'FormulaCode'}, $headingb); $wrksheet->write_string( 'G8', 'Customer Name:', $heading); $wrksheet->merge_range_type( 'string', 'I8:M8', $form_obj->{'CustomerName'}, $headingb); 
    $wrksheet->write_blank( 8, 0, $heading);
    $wrksheet->merge_range_type( 'string', 'A10:B10', 'Product Name:', $heading); $wrksheet->merge_range_type( 'string', 'C10:E10', $ppro_obj->{'Description'}, $headingb); $wrksheet->write_string( 'G10', 'Lot Number:', $heading); $wrksheet->write_string( 'I10', $batch_obj->{'Lot'}, $headingb); $wrksheet->write_string( 'K10', 'Mfg Date:', $heading); 
    if ( $end_date[0] > 2000 ) {
        $wrksheet->merge_range_type( 'string', 'L10:M10', $close_date, $headingb); 
    } else {
        $wrksheet->merge_range_type( 'string', 'L10:M10', "Pending", $headingb); 
    }
    $wrksheet->merge_range_type( 'string', 'A11:N11', q{}, $bborder);
    $wrksheet->set_row( 11,8 );
    $wrksheet->write_blank( 11, 0, $body);    
    $wrksheet->merge_range_type( 'string', 'A13:J13', "Every (".$form_obj->{'ServingQty'}.") $dose contain(s) the following:" , $body);
    $wrksheet->set_row( 13,8 );
    $wrksheet->write_blank( 13, 0, $body);    
    $wrksheet->write_string( 'A15', "ID Number", $bodyb); $wrksheet->merge_range_type( 'string', 'B15:E15', "Component Name", $bodyb); $wrksheet->merge_range_type( 'string', 'F15:G15', "Analyte", $bodyb ); $wrksheet->write_string( 'H15', "Ovg\%", $bodycb); $wrksheet->write_string( 'I15', "Claim", $bodybr); $wrksheet->write_string( 'J15', "(".$form_obj->{'ServingQty'}.") ", $bodyb); $wrksheet->write_string( 'K15', "Formulation", $bodybr); $wrksheet->write_string( 'L15', "(1) ", $bodyb);
    $wrksheet->write_string( 'M15', "Production" , $bodybr); $wrksheet->write_string( 'N15', '(Kg)' , $bodyb);
    $wrksheet->write_string( 'A16', "------------" , $body); $wrksheet->merge_range_type( 'string', 'B16:E16', "----------------------------------------------------------" , $body); $wrksheet->merge_range_type( 'string', 'F16:G16', "----------------------------------" , $body); $wrksheet->write_string( 'H16', "----" , $bodyc); $wrksheet->write_string( 'I16', "------------" , $bodyr); $wrksheet->write_string( 'J16', "---" , $body); $wrksheet->write_string( 'K16', "------------" , $bodyr); $wrksheet->write_string( 'L16', "---" , $body);
    $wrksheet->write_string( 'M16', "------------" , $bodyr); $wrksheet->write_string( 'N16', '---' , $body);

    my @second_list;
    my @sorted_list = sort { $a->[5] <=> $b->[5] } @list;
    foreach my $count ( 0 .. $#sorted_list ) {
        if ( $sorted_list[$count][3] =~ m/other/ ) {
            push @second_list, $sorted_list[$count];
        } else { unshift @second_list, $sorted_list[$count]; }
    }

    my $row = 17;
    foreach my $count ( 0 .. $#second_list +1) {   
        if ($row % 42 == 0) {
            #$wrksheet->set_h_pagebreaks( $row-1 );
            $wrksheet->write_string( $row-1,0, "------------" , $body); $wrksheet->merge_range_type( 'string', $row-1,1,$row-1,4, "----------------------------------------------------------" , $body); $wrksheet->merge_range_type( 'string', $row-1,5,$row-1,6, "----------------------------------" , $body); $wrksheet->write_string( $row-1,7, "----" , $bodyc); $wrksheet->write_string( $row-1,8, "------------" , $bodyr); $wrksheet->write_string( $row-1,9, "---" , $body); $wrksheet->write_string( $row-1,10, "------------" , $bodyr); $wrksheet->write_string( $row-1,11, "---" , $body);
            $wrksheet->write_string( $row-1,12, "------------" , $bodyr); $wrksheet->write_string( $row-1,13, '---' , $body);
            $row = $row+1;
        }
        $wrksheet->write_string( $row-1,0, "$second_list[$count][0]" , $body); $wrksheet->merge_range_type( 'string', "B$row:E$row", "$second_list[$count][1]" , $body); $wrksheet->merge_range_type( 'string', "F$row:G$row", "$second_list[$count][2]" , $body); $wrksheet->write_string( $row-1,7, "$second_list[$count][7]" , $bodyc); $wrksheet->write_string( $row-1,8, "$second_list[$count][3]" , $bodyr); $wrksheet->write_string( $row-1,9, "$second_list[$count][4]" , $body); $wrksheet->write_string( $row-1,10, "$second_list[$count][5]" , $bodyr); $wrksheet->write_string( $row-1,11, "$second_list[$count][6]" , $body);
        $wrksheet->write_string( $row-1,12, "$second_list[$count][8]" , $bodyr); $wrksheet->write_string( $row-1,13, "$second_list[$count][9]" , $body);
        $row++;
    }

    $wrksheet->write_string( $row, 1, "Total", $bodyb); $wrksheet->write_string( $row, 10, "$ftotal", $bodyr); $wrksheet->write_string( $row, 11, "mg", $body);
     $wrksheet->write_string( $row, 12, "$ptotal", $bodyr); $wrksheet->write_string( $row, 13, "Kg", $body);

    $row=$row+2;
    $wrksheet->merge_range_type('string', "A$row:J$row", "Allergens: *".$ppro_obj->{'AllergenList'}, $bodyb);

    $row=$row+2;
    $wrksheet->set_row($row-1, 150);
    $wrksheet->merge_range_type('string', "A$row:L$row", "Method of Analysis - Input data, data derived from each qualified ingredient\'s Certificate of Analysis and MMR. Additionally, at least two raw ingredients are quantitatively analyzed using a sound statistically valid methodology. FTIR Match, Heavy Metal Analysis and Microbial Testing are also performed on this lot. These results are posted to your account at http:\\www.columbianutritional.com.\n
* FDA Food Facility Registration Number 13657096712\n
* Bioterrorism Act of 2003 registered facility\n", $parag);
#* Certified by NSF International to be a GMP (Good Manufacturing Practices) compliant facility, as prescribed by 21 CFR Section 111 which was published by the FDA in June 2007\n
#* NSF Certification Number 3E561-02

    $row = $row+2;
    $wrksheet->merge_range_type('string', "A$row:E$row", "Quality Assurance", $sigline); $wrksheet->merge_range_type('string', "G$row:H$row", "Date", $sigline);
    $wrksheet->set_row( $row++, 20); 
    $wrksheet->write_blank( $row, 0, $sigline);  
    $row++;
    $wrksheet->merge_range_type('string', "A$row:E$row", "______________________________", $sigline); $wrksheet->merge_range_type('string', "G$row:H$row", "_______________", $sigline);
    
    if ($row < 50) {
            $wrksheet->fit_to_pages( 1,1);
    }
    return ( 1 );
}


sub make_preblend {
#####
    #
    # Make Preblend instructions
    #
#####
    local $| = 1;
    my @args    = @_;
    my $product = $args[0];
    my $ppro_form  = PProFormula->new($product);
    my ( $batchnumber, $index, $threshold );    

    $log->log( ">>>make_preblend $product", 2 );
    while (1) {
        my (%list);
        print "\nBR and LOT number(s) in ProcessPro (only the past year is displayed):\n";
        print "\n        BR\tLOT NO  \tMFG DATE\n";
        foreach my $count ( 0 .. $#{ $ppro_form->{'Wono'} } ) {
            my @dt       = split( q{ }, $ppro_form->{'Wono'}[$count][3] );
            my @end_date = split( q{-}, $dt[0] );
            my $mfg_date;
            if (   ( $end_date[0] < 2000 )
                && ( $ppro_form->{'Wono'}[$count][4] == 5 ) ) {
                $mfg_date = "- Voided";
            } elsif ( $ppro_form->{'Wono'}[$count][4] == 4 ) {
                $mfg_date = sprintf "%02d/%02d/%04d", $end_date[1], $end_date[2], $end_date[0];
            } else {
                $mfg_date = "- Open";
            }
            $list{ int( $ppro_form->{'Wono'}[$count][1] ) } = $count;
            if (( $end_date[0] > ($date[5] + 1898) )||( $end_date[0] < 2000 )) {
                printf "%s\t%s\t%s\n", $ppro_form->{'Wono'}[$count][1], $ppro_form->{'Wono'}[$count][2], $mfg_date;
            }
        }

        $batchnumber = prompt("\tEnter Batch Record to use <x - cancel>: ");
        
        if ($batchnumber =~ m/^x$/i) {
            $log->log(">> Document canceled!",1);
            return 0;
        } elsif ( (!exists $list{$batchnumber}) || ( int($ppro_form->{'Wono'}[$list{$batchnumber}][2]) !~ m/^\d+$/) ) {
            print "\n\t$batchnumber is invalid!\n";
            next;
        }
        $index = $list{$batchnumber};
        last;
    }

    if ( $ppro_form->{'Wono'}[$index][4] < 4 ) {
        $log->append("WARNING $bell: This BR has not been closed. Changes may still occur!", 1 );
    }

    my $batch = Batch->new($batchnumber);
    $log->log( " > BR used:\t$batch->{'BatchRecord'}",         1 );
    $log->log( " > Total Allocated:\t$batch->{'AllocQty'} Kg", 1 );

    if (!$batch->{'Yield'}) {
        $log->log( "!! Unable to pull data for $batch->{'FormulaCode'} v$batch->{'Version'} BR $batch->{'BatchRecord'} ",1); return 0;
    }

    my $runsize = sprintf "%.4f", $batch->{'RouteSize'} / $batch->{'Yield'} ;
    my $runqty = int($batch->{'AllocQty'} / $runsize);
    my $mod = ( $batch->{'AllocQty'} / $runsize ) - $runqty ;
    my $runrmd = sprintf "%.4f", $mod * $runsize;

    while (1) {
        my $input = prompt("\tEnter the Raw Material threshold <default - 1 Kg, x - cancel>: " );
        if (!$input) {
            $threshold = 1.000;
            last;
        } elsif ( $input =~ m/^x$/i ) {
            $log->log(">> Document canceled!",1);
            return;
        } elsif ( $input !~ m/\d+/ ) {
            print "\n\t$input Kg is an invalid threshold!\n";
            next;
        } else {
            $threshold = sprintf "%.3f",$input;
            last;
        }
    }
    $log->log( " > Preblend threshold:\t" . $threshold . " Kg", 1 );
    
    if ( $batch->{'RouteSize'} > 1 ) {        
        $log->log( " > Runs:\t" . $runqty . " @ \t$runsize Kg", 1 );
        $log->log( " > Remainder Size:\t$runrmd Kg", 1 );
        while (1) {
            print "\n***\tThis is a multi-run batch.\n\n";
            my $input = prompt("\tDo you want to print the \n\t<W>hole run, <A>ll runs, run #<1-" . ($mod>0?$runqty+1:$runqty) . ">, or <X> - cancel ?" );
            if ( $input =~ m/^x$/i ) {
                $log->log(">> Document canceled!",1);
                last;
            } elsif ( $input =~ m/^a$/i ) {
                foreach my $count (1..$runqty) {
                    write_preblend( $ppro_form, $batch, $runsize, $count, $threshold  );
                    next;
                }
                if ( $mod > 0 ) {
                    write_preblend( $ppro_form, $batch, $runrmd, $runqty+1, $threshold  );
                    last;
                }
                last;
            } elsif ( $input =~ m/^w$/i ) {
                write_preblend( $ppro_form, $batch, $batch->{'AllocQty'}, 0, $threshold  );
                last;
            } elsif ( ($input <= $runqty ) && ($input > 0) ) {
                write_preblend( $ppro_form, $batch, $runsize, $input, $threshold  );
                last;
            } elsif ( ($input > $runqty ) && ($runrmd != 0) ) {
                write_preblend( $ppro_form, $batch, $runrmd, $runqty+1, $threshold  );
                last;
            } else {
                print "\n\t$input is invalid!\n";
                next;
            }
        }
    } else {
        write_preblend( $ppro_form, $batch, $batch->{'AllocQty'}, 0, $threshold  );
    }
}

sub write_preblend {
#####
    #
    # Make Preblend instructions
    #
#####
    local $| = 1;
    my @args    = @_;    
    my $ppro_obj  = shift @args;
    my $batch_obj = shift @args;
    my $runsize   = shift @args;
    my $run       = shift @args;
    my $threshold = shift @args;

    my $preblend  = 0;
    my $pbcount   = 0;
    my $rawcount  = $#{ $ppro_obj->{'Details'} };
    my $pbtheo    = $runsize * .05;      

    my ( @html, $complete );

    my $theo_yield;

# # cover_pg1
    $log->log( " > Printing PreBlend Pg1...", 1 );
    push @html, '
<!DOCTYPE HTML>
<html><head>
<title>PreBlend for ' . $batch_obj->{'FormulaCode'} . ' Lot Number ' . $batch_obj->{'Lot'} . ( $run ? "\." . $run : q{} ) . ' ' . ($threshold != 1 ? $threshold . ' Kg threshold':q{}) .'</title>
<style><!--
    @page {
        size: letter landscape;
        margin: 7px 40px 2px 40px;
        padding: 0px; }
    @media print {
        .pagebody {  }
        .noprint { display: none; }
    }
    body {
        counter-reset: page 0;
        font-size:11pt; }
    table	{
        width: 100%;
        height: 8in;
        border: 0px solid black;
        vertical-align: middle; }
    td	{
        height: 15px;
        width: 1vw;
        border: 0px solid black;
        padding-top:1px;
        padding-right:1px;
        padding-left:1px;
        color:windowtext;
        font-size:11pt;
        font-weight:350;
        font-style:normal;
        text-decoration:none;
        font-family:Verdana,Helvetica;
        text-align:general;
        vertical-align:bottom;}
    th	{
        height: 15px;
        width: 1vw;
        border: 0px solid black;
        padding-top:1px;
        padding-right:1px;
        padding-left:1px;
        color:windowtext;
        font-size:12pt;
        font-family:Verdana,Helvetica;
        text-align:general;
        vertical-align:bottom;} 
    li  {
        margin-left: 21px;
        text-indent: -21px;
        font-size:11pt;
        list-style-type: none;}

    .list {
        counter-reset: itemlist;}
    .list li:before {
        content: counter(itemlist, lower-alpha) ") ";
        counter-increment: itemlist;} 
    .heavy	{
        color:black;
        font-size:12pt;
        font-weight:700;
        font-family:Verdana,Helvetica;
        vertical-align: bottom;
        white-space: nowrap; }
    .QA {
        outline: thin solid black; }
    .header {
        top: 0px;
        height: 22px;
        padding-top: 10px;
        font-size: 9pt;
        text-align: center;    }
    .footer {
        bottom: 0px;
        height: 27px;
        padding-top: 0px;
        font-size: 9pt;
        text-align: right; }
    .header:after {
        counter-increment: page;
        content: "Preblend for ' . $batch_obj->{'FormulaCode'} . ($ppro_obj->{'Revision'} > 0? ' rev' . sprintf "%02d", $ppro_obj->{'Revision'} : q{} ) . ' Lot Number ' . $batch_obj->{'Lot'} . ( $run ? "\." . $run : q{} ) . ' (' . $pbtheo . ' Kg) ' . ($threshold != 1 ? $threshold . ' Kg threshold':q{})  .'"; }
    .footer:after {
        content: "v' .$VERSION. ' printed ' . $day . ' " " page " counter(page); }
--></style>';
    push @html, '</head><body>';
    
    foreach my $count ( 0 .. $rawcount ) {
        my $portion = $ppro_obj->{'Details'}[$count][2] * $runsize;
        if ( $portion <= $threshold ) { 
            $preblend += $portion; 
            $pbcount++;}
    }
    $log->log( " > Preblend items: $pbcount", 1 );    
    $log->log( " > Preblend portion: $preblend Kg", 1 );

    if ($pbtheo < $preblend) {
        $pbtheo = $preblend;
    }

# # pb_pg3
    $log->log( " > Printing Pre-blend ...", 1 );
    push @html, $main_tableh;
    if (($rawcount-$pbcount>0)&&($pbcount>0)) { 
        push @html, '
     <tr>
      <td colspan="91" class="heavy">Pre-Blend Instructions<img src="http://icons.iconarchive.com/icons/icojam/onebit-4/32/printer-icon.png" width="30" onclick="window.print();" style="cursor:pointer;" class="noprint"></td>
      <td colspan="9" align="right">MF-214</td>
     </tr>';   
        push @html, '
     <tr style="height:46px">
        <td colspan="100">
            <li style="list-style: disc inside">All ingredients weighing 1 Kg or less, are to be added to the Pre-Blend.</li>
            <li style="list-style: disc inside">Ingredients weighing 100 gm or less, must be wieghed using the Gram scale.</li>
            <li style="list-style: disc inside">Total Pre-Blend weight must weigh at least 5% of total blend weight</li></td>
     </tr>';
        push @html, $single_blank;
#5 lines
        push @html, '
     <tr>
      <td colspan="100" class="heavy">Total Pre-Blend Weight (Theoretical: ' . (sprintf "%.3f", $pbtheo) . '&nbsp;Kg)</td>
     </tr>';
        push @html, $single_blank;
        push @html, '
     <tr>
      <td colspan="100" class="heavy">Add the following ingredients to the Pre-blend drum:</td>
     </tr>';
        push @html, $single_blank; 
        push @html, '
     <tr>
       <td colspan="4">&nbsp;</td>
       <td colspan="9" class="heavy">Code</td>
       <td colspan="30" class="heavy">Description</td>
       <td colspan="15" class="heavy" align="center">Batch Qty</td>
       <td colspan="42">&nbsp;</td>
     </tr>';    
#5 lines
        my ($preb,$macroid);
        my ($pagec,$linec,$macropct,$paget)=0;
        $paget = int(($pbcount - 13) / 15)+1;
        foreach my $count ( 0 .. $#{ $ppro_obj->{'Details'} } ) {
            my $pctg = $ppro_obj->{'Details'}[$count][2];
            if ($pctg > $macropct) { # Determine largest (macro) raw material
                $macropct = $pctg;
                $macroid = $ppro_obj->{'Details'}[$count][0];
            }
            my $portion = $pctg * $runsize;
            if ( sprintf("%.3f",$portion) <= $threshold ) {
                my $curr_raw = PProRaw->new($ppro_obj->{'Details'}[$count][0]);     
                $linec++;
                push @html,'
     <tr style="height:42px">
       <td colspan="4">&nbsp;</td>
       <td colspan="9" style="vertical-align:middle">' . $curr_raw->{'ItemCode'} . '</td>
       <td colspan="30" style="vertical-align:middle">' . $curr_raw->{'Description'} . '</td>
       <td colspan="15" align="center" style="border:1px solid lightgray">' . sprintf("%.3f&nbsp;Kg",$portion) . (($portion<0.5)? sprintf("<br>(%.3f&nbsp;gm)",$portion*1000):q{}) . '</td>
       <td colspan="11" align="center">Date:&nbsp;_________</td>
       <td colspan="11" align="center">Time:&nbsp;________<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
       <td colspan="11" align="center">Initials:&nbsp;_________</td>
       <td colspan="9">&nbsp;</td>
     </tr>';
            }
            if ((($linec>15)&&($pagec<$paget)) || (($linec>11)&&($pagec<1)) || (($linec>10)&&($pagec==$paget))) { 
                push @html, $single_blank;
                push @html, '
     <tr>
       <td colspan="100" align="right">(Continued on next page)</td>
     </tr>';
                push @html, $single_blank; 
                push @html, $main_tablef;
                $pagec++; $linec=0;
                push @html, $main_tableh;
                push @html, '
     <tr>
       <td colspan="100" align="right">(Continued from previous page)</td>
     </tr>';
                push @html, '
     <tr>
      <td colspan="100" class="heavy">Add the following ingredients to the Pre-blend drum:</td>
     </tr>';
                push @html, $single_blank; 
                push @html, '
     <tr>
       <td colspan="4">&nbsp;</td>
       <td colspan="9" class="heavy">Code</td>
       <td colspan="30" class="heavy">Description</td>
       <td colspan="15" class="heavy" align="center">Batch Qty</td>
       <td colspan="42">&nbsp;</td>
     </tr>';
            }
        }

        my $macroraw = PProRaw->new($macroid);
        my $macroweight = $pbtheo-$preblend;
        my $macrotheo = $macropct * $runsize;
        my $macroremain = $macrotheo - $macroweight;
        if ($macroweight) {
            $linec += 3;
            push @html, $single_blank;
            push @html, '
         <tr>
           <td colspan="100" class="heavy">Add the following Macro ingredient to Pre-blend:</td>
         </tr>';
            push @html, $single_blank;    
            push @html,'
         <tr>
           <td colspan="4">&nbsp;</td>
           <td colspan="9" style="vertical-align:middle">' . $macroraw->{'ItemCode'} . '</td>
           <td colspan="30" style="vertical-align:middle">' . $macroraw->{'Description'} . '</td>
           <td colspan="15" align="center" style="border:1px solid lightgray">' . sprintf("%.3f&nbsp;Kg",$macroweight) . ($macroweight<.5?sprintf("&nbsp;(%.3f&nbsp;gm)",$macroweight*1000):q{}) . '</td>
           <td colspan="11" align="center">Date:&nbsp;_________</td>
           <td colspan="11" align="center">Time:&nbsp;________<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
           <td colspan="11" align="center">Initials:&nbsp;_________</td>
           <td colspan="9">&nbsp;</td>
         </tr>';
            push @html, $single_blank; 
            push @html,'
         <tr>
           <td colspan="13">&nbsp;</td>   
           <td colspan="87"><font size=-1>Remainder of ' . $macroraw->{'ItemCode'} . '&nbsp;<span style="color:black;font-weight:700;white-space:wrap;">(' . sprintf("%.3f&nbsp;Kg",$macroremain) . ($macroremain<0.5 ? sprintf("&nbsp;(%.3f&nbsp;gm)",$macroremain*1000):q{}) . ')</span> goes into Main Blend.</font></td>
         </tr>';
        } else {
            $linec += 1;
            push @html, $single_blank; 
            push @html,'
         <tr>
           <td colspan="13">&nbsp;</td>   
           <td colspan="87"><font size=-1>No Macro ingredient to add into Main Blend</font></td>
         </tr>';
        }
        for (my $c=0; $c<10-$linec; $c++) {
            push @html, $single_blank; }
        push @html, $single_blank;
        push @html, '
     <tr>
      <td colspan="91" align="left" class="heavy">Pre-blend Particle Screening</td>
      <td colspan="9" align="right">MF-214</td>
     </tr>';
        push @html, '
     <tr>
      <td colspan="91"><li style="list-style: disc inside">Screen the Pre-blend through the mesh, and keep separate from the Main Blend.</li></td>
      <td colspan="9">&nbsp;</td>
     </tr>';
        push @html, $single_blank;
        push @html, '
     <tr>
      <td colspan="25">Mesh Used:</td>
      <td colspan="12" align="center">Mesh&nbsp;#_________</td>
      <td colspan="12" align="center">Size&nbsp;_________</td>
      <td colspan="9">&nbsp;</td>
      <td colspan="11" align="center">Date:&nbsp;_________</td>
      <td colspan="11" align="center">Time:&nbsp;________<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
      <td colspan="11" align="center">Initials:&nbsp;_________</td>
      <td colspan="9">&nbsp;</td>
     </tr>';
        push @html, $single_blank;
        push @html, '
     <tr>
       <td colspan="25" class="heavy">Pre-blend Total Weight:</td>
       <td colspan="12" class="heavy" align="center">_________&nbsp;Kg</td>
       <td colspan="21">&nbsp;</td>
       <td colspan="11" align="center">Date:&nbsp;_________</td>
       <td colspan="11" align="center">Time:&nbsp;________<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
       <td colspan="11" align="center">Initials:&nbsp;_________</td>
       <td colspan="9">&nbsp;</td>   
     </tr>';
        push @html, $single_blank;
        push @html, '
     <tr>
      <td colspan="25">Is Pre-blend at least 5% of total blend weight?</td>
      <td colspan="6" align="center">Yes</td>
      <td colspan="6" align="center">No</td>
      <td colspan="21">&nbsp;</td>
      <td colspan="11" align="center">Date:&nbsp;_________</td>
      <td colspan="11" align="center">Time:&nbsp;________<span style="font-weight: 350; font-size: 10px;">&nbsp;am/pm</span></td>
      <td colspan="11" align="center">Initials:&nbsp;_________</td>
      <td colspan="9">&nbsp;</td>
     </tr>';
        push @html, $single_blank; 
        push @html, $main_tablef;

        my $path = "$ENV{USERPROFILE}\\Desktop\\";
        my $file = sprintf "%d%s %s %s %s.html", $batch_obj->{'Lot'}, ( $run ? q{.} . $run : q{} ), $ppro_obj->{'FormulaCode'}, $ppro_obj->{'Description'}, $pbtheo . ' Kg Pre-Blend';
        $file =~ s|\#||g;
        $file =~ s|/|_|g;
        $file =~ s|^a-zA-Z\d\s[()+,-]|_|g;
        my $doc = LogFile->new( $file, $path );
        $doc->write( "@html", 2 );
        $log->log( ">> $doc->{'FilePath'} completed!", 1 );
    } else {
        $log->log( ">> No Pre-blend!", 1 );
        print "\nNo Pre-Blend to generate!\n";
    }
    return 1;
}





sub MainMenu {
#####
    #
    # MainMenu
    #
    #
#####
    while (1) {
        print "\n\n";
        print "Main Menu -\n\n";
        print "\t1\t- Xfer Formula or Items between DBs\n";
        print "\t2\t- Generate production records\n";
        print "\t3\t- Map Formula ID across DBs\n"; 

        if ($username eq "CTRICHEL") { print "\t4\t- DB Object Dump\n"; }
        print "\n\tx\t- Exit\n";

        my $input = prompt("\tCommand : ");

        if ( $input =~ m/^x$/i ) {
            last;
        } elsif ( $input == 1 ) {
            menuXfer();
        } elsif ( $input == 2 ) {
            menuHTML();
        } elsif ( $input == 3 ) {
            menuMapFormula();
        } elsif (( $input == 4 ) && ($username eq "CTRICHEL")) {
            menuMapObj();
        }
        next;
    }
    return 1;
}

sub menuMapObj {
#####
    #
    # Interface for mapping Formula across DBs
    #
#####
    while (1) {
        print "\n\n";
        print "Dump DB Objects\n\n";
        print "\t<FormulaID>\t- Dump objects associated with <FormulaID>\n";
        print "\n";
        print "\tm\t\t- Main Menu\n";

        my $input = uc prompt("\tEnter Formula ID : ");

        if ( $input =~ m/^m$/i ) {
            return;
        } else {
            dump_obj($input);
        }
        next;
    }
    return (1);
}

sub menuMapFormula {
#####
    #
    # Interface for mapping Formula across DBs
    #
#####
    while (1) {
        print "\n\n";
        print "Map Formula ID across DBs\n\n";
        print "\t<FormulaID>\t- Find instances of <FormulaID> across Databases\n";
        print "\n";
        print "\tm\t\t- Main Menu\n";

        my $input = uc prompt("\tEnter Formula ID : ");

        if ( $input =~ m/^m$/i ) {
            return;
        } else {
            locate_db($input);
        }
        next;
    }
    return (1);
}

sub menuXfer {
#####
    #
    # Interface for transfering between systems
    #
#####
    my ( $input, $item );
    while (1) {
        print "\n\n";
        print "Transfer between DBs -\n\n";
        print "\t1\t- Transfer Formula to ProcessPro\n";
        print "\t2\t- Transfer Item to ProcessPro\n";
        print "\n";
        print "\t3\t- Transfer Formula to Formulator\n";
        print "\t4\t- Transfer Item to Formulator\n";
        print "\t5\t- Transfer Packaging to Formulator\n";
        print "\n";
        print "\tm\t- Main Menu\n";

        $input = uc prompt("\tEnter : ");

        if ( $input =~ m/^m$/i ) {
            return;
        } elsif ( $input == 1 ) {
            $item = uc prompt("\tEnter the Formula to transfer : ");
            xfer_form2p($item);
        } elsif ( $input == 2 ) {
            $item = uc prompt("\tEnter the Item Number to transfer : ");
            xfer_item2p( $item, 0 );
        } elsif ( $input == 3 ) {
            $item = uc prompt("\tEnter the Formula to transfer : ");
            xfer_form2f($item);
        } elsif ( $input == 4 ) {
            $item = uc prompt("\tEnter the Item Number to transfer : ");
            xfer_item2f( $item, 0 );
        } elsif ( $input == 5 ) {
            $item = uc prompt("\tEnter the Item Number to transfer : ");
            xfer_pkg2f( $item, 0 );
        }
        next;
    }

    return 1;
}

sub menuHTML {
#####
    #
    # Interface for generating MMR
    #
#####
    my ( $input, $item );
    while (1) {
        print "\n\n";
        print "Create Production Documents -\n\n";
        print "\t1\t- Create MMR file\n";
        print "\t2\t- Create BR file\n";
        print "\t3\t- Create CofM file\n";
        print "\t4\t- Create PreBlend file\n";
        print "\n";
        print "\tm\t- Main Menu\n";

        $input = uc prompt("\tEnter : ");

        if ( $input =~ m/^m$/i ) {
            return;
        } elsif ( $input == 1 ) {
            $item = uc prompt("\tEnter the Formula to make an MMR for : ");
            if ( !FormulaTool->validate_formula($item) ) { make_mmr( $item, 1 ); }
        } elsif ( $input == 2 ) {
            $item = uc prompt("\tEnter the Formula to make an BR for : ");
            if ( !FormulaTool->validate_formula($item) ) { make_mmr( $item, 0 ); }
        } elsif ( $input == 3 ) {
            $item = uc prompt("\tEnter the Formula to make an CofM for : ");
            if ( !FormulaTool->validate_formula($item) ) { make_cofm($item); }
        } elsif ( $input == 4 ) {
            $item = uc prompt("\tEnter the Formula to make a PreBlend for : ");
            if ( !FormulaTool->validate_formula($item) ) { make_preblend($item); }
        } 
        next;
    }
    return 1;
}

sub prompt {
#####
    # prompt(text) -
    #	Displays text and returns keyboard input as string
    #
#####
    my @args = @_;
    my $text = $args[0];
    print "\n$text";
    my $input = <>;
    chomp $input;
    print "\n";
    $input =~ s/[\s\!\?\/\\\<\>\,\'\"\@\(\)\[\]]//g;
    return $input;
}
