package FormulaTool;
use strict;
our ($VERSION) = '1.6.3';

$SIG{__WARN__} = sub { warn sprintf("[%s] ", scalar localtime), @_ };
$SIG{__DIE__}  = sub { die  sprintf("[%s] ", scalar localtime), @_ };

use DBI;
use DBD::ODBC;
our $parentfile = (caller(0))[1];

our $formdb_user = q{};
our $formdb_pass = q{};
#our $formdb_name = 'dbi:ADO:Provider=Microsoft.ACE.OLEDB.12.0;Data Source= \\\Appserver\Formulator\Data\CNS - Repaired.mdb';
our $formdb_name = 'dbi:ADO:Provider=Microsoft.Jet.OLEDB.4.0;Data Source= \\\Appserver\Formulator\Data\CNS - Repaired.mdb';
our $specdb_user = q{};
our $specdb_pass = q{};
our $specdb_name = 'dbi:ADO:Provider=Microsoft.ACE.OLEDB.12.0;Data Source= \\\Appserver\company_share\Specifications\Specs.accdb';
#our $specdb_name = 'dbi:ADO:Provider=Microsoft.ACE.OLEDB.12.0;Data Source= \\Appserver\company_share\Specifications\Specs.accdb';
our $pprodb_name;
$pprodb_name = 'dbi:ODBC:driver={SQL Server Native Client 10.0};Server=SQLSERVER; database=db01;uid=ppsa;pwd=fufi2*'; #Live
print "live_env\n";
#our $pprodb_name = 'dbi:ODBC:driver={SQL Server Native Client 10.0};Server=ADSERVER; database=rd;uid=ppsa;pwd=fufi2*'; #R&D
#our $pprodb_name = 'dbi:ODBC:driver={SQL Server Native Client 10.0};Server=ADSERVER; database=testrddb;uid=ppsa;pwd=fufi2*'; #TestR&D
#our $pprodb_name = 'dbi:ODBC:driver={SQL Server Native Client 10.0};Server=ADSERVER; database=testdb;uid=ppsa;pwd=fufi2*'; #Test


sub connect_formdb {
    my $dbh = DBI->connect_cached( $formdb_name, { 
  PrintError => 1, # warn() on errors 
  RaiseError => 0, # don't die on error 
  AutoCommit => 1, # commit executes immediately 
 } ) or die "!!!Could not connect to $formdb_name\n\t$DBI::errstr";
    return $dbh;
}

sub connect_pprodb {
    my $dbh = DBI->connect_cached( $pprodb_name, { 
  PrintError => 1, # warn() on errors 
  RaiseError => 0, # don't die on error 
  AutoCommit => 1, # commit executes immediately 
 } ) or die "!!!Could not connect to $pprodb_name\n\t$DBI::errstr";
    return $dbh;
}

sub connect_specdb {
    my $dbh = DBI->connect_cached( $specdb_name, { 
  PrintError => 1, # warn() on errors 
  RaiseError => 0, # don't die on error 
  AutoCommit => 1, # commit executes immediately 
 } ) or die "!!!Could not connect to $specdb_name\n\t$DBI::errstr";
    return $dbh;
}

sub querydb {
# querydb(dbh,selection,table,<filter>,<echo>) -
#	Query table in db for selection using <filter>, <echo> flags output
#
    my $class     = shift;
    my $dbh       = shift;
    my @args = @_;
    my $selection = $args[0];
    my $table     = $args[1];
    my $filter    = $args[2];
    my $statement =
      'select ' . $selection . ' from ' . $table . ' where ' . $filter;

    my $sth = $dbh->prepare($statement)
      or die(" !!Could not prepare [$statement] \nerror: $!");

    $sth->execute
      or die(" !!Could not execute [$statement] \nerror: $!");

    my @result = $sth->fetchall_arrayref;
    $sth->finish;

    return @result;
}

sub validate_formula {
# Verify BOM exists in Process Pro
#
    my $class = shift;
    my @args = @_;
    my $id = FormulaTool->despace($args[0]);
    my $valid = 3;
    my $pprodb = connect_pprodb;
    my @presult = FormulaTool->querydb($pprodb,'item,revstat,inactive','ppbmhd',"item like '$id'");
    my $formdb = connect_formdb;
    my @fresult = FormulaTool->querydb($formdb,'FormulaCode,Status','FormulaMaster',"FormulaCode like '$id'");
    
    if ( $presult[0][0][1] ) { $valid = $valid - 1; }
    if ( $fresult[0][0][0] ) { $valid = $valid - 2; }
 
    if ($valid == 0) { 
    	print " > BOM $id is valid in PPro and Formulator\n";
    } elsif ($valid == 1) {
    	print " > BOM $id is NOT in PPro\n";
    } elsif ($valid == 2) {
        print " > BOM $id is NOT in Formulator\n";
    } else {
        print " > BOM $id is INVALID in both locations\n";
    }
    
    $pprodb->disconnect;
    $formdb->disconnect;
    return $valid;
}

sub validate_raw {
# Verify BOM exists in Process Pro
#
    my $class = shift;
    my @args = @_;
    my $id = FormulaTool->despace($args[0]);
    my $valid = 3;
    my $pprodb = connect_pprodb;
    my @presult = FormulaTool->querydb($pprodb,'item,itemstat','icitem',"item like '$id'");
    my $formdb = connect_formdb;
    my @fresult = FormulaTool->querydb($formdb,'ItemCode,Status','RawMaterials',"ItemCode like '$id'");

    if ( $presult[0][0][1] ) { $valid = $valid - 1; }
    if ( $fresult[0][0][0] ) { $valid = $valid - 2; }
 
    if ($valid == 0) { 
    	print " > Raw Material $id is valid in PPro and Formulator\n";
    } elsif ($valid == 1) {
    	print " > Raw Material $id is NOT in PPro\n";
    } elsif ($valid == 2) {
        print " > Raw Material $id is NOT in Formulator\n";
    } else {
        print " > Raw Material $id is INVALID in both locations\n";
    }
    
    $pprodb->disconnect;
    $formdb->disconnect;
    return $valid;
}

sub validate_pkg {
# Verify pkg exists in Process Pro
#
    my $class = shift;
    my @args = @_;
    my $id = FormulaTool->despace($args[0]);
    my $valid = 3;
    my $pprodb = connect_pprodb;
    my @presult = FormulaTool->querydb($pprodb,'item,itemstat','icitem',"item like '$id'");
    my $formdb = connect_formdb;
    my @fresult = FormulaTool->querydb($formdb,'InstrCode,Status','BOMPackaging',"InstrCode like '$id'");

    if ( $presult[0][0][1] ) { $valid = $valid - 1; }
    if ( $fresult[0][0][0] ) { $valid = $valid - 2; }
 
    if ($valid == 0) { 
    	print " > Packaging $id is valid in PPro and Formulator\n";
    } elsif ($valid == 1) {
    	print " > Packaging $id is NOT in PPro\n";
    } elsif ($valid == 2) {
        print " > Packaging $id is NOT in Formulator\n";
    } else {
        print " > Packaging $id is INVALID in both locations\n";
    }
    
    $pprodb->disconnect;
    $formdb->disconnect;
    return $valid;
}

sub despace {
# Remove spaces from string
#
    my $class = shift;
    my @args = @_;
    my $string = $args[0];

    $string =~ s/^\s+|\s+$//sg;
    $string =~ s/'//sg;
    $string =~ s/\s\s+/ /sg;

    return $string;
}

sub vardump {
# Use data dumper to dump variable
#
    use Data::Dumper;
    $Data::Dumper::Indent = 2;
    $Data::Dumper::Useqq  = 0;
    $Data::Dumper::Terse  = 1;
    my $class = shift;
    my $v;
    
    $v = Data::Dumper->Dump(\@_);
    print $v;
    return $v;
}
1;