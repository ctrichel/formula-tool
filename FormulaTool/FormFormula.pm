package FormFormula;

use strict;
our $VERSION = '1.1.3';

use FormulaTool;

sub new {
# Formulator Formula
#  	FormulaCode
# 	FormulaMaster.Description   => 'Description'
# 	FormulaMaster.Class         => 'Class'
# 	FormulaMaster.SubClass      => 'SubClass'
# 	FormulaMaster.ServingSize   => 'ServingWeight'
# 	FormulaMaster.ServingUOM    => 'ServingUOM'
# 	FormulaMaster.Status        => 'Status'
# 	FormulaMaster.Memo01        => 'ServingDesc'
# 	FormulaMaster.Memo02        => 'ServingSize'
# 	FormulaMaster.Memo03        => 'ServingQty'
# 	FormulaMaster.Memo05        => 'Appearance'
# 	FormulaMaster.AddedBy       => 'AddedBy'
# 	FormulaMaster.Added         => 'Added'
# 	FormulaMaster.DefaultUOM    => 'DefaultUOM'
# 	FormulaMaster.FormulaTotal  => 'FormulaTotal'
# 	FormulaMaster.FormulaUOM    => 'FormulaUOM'
# 	FormulaMaster.CustomerCode  => 'CustomerCode'
# 	FormulaDetail.LineNumber    => 'RawMaterials'[0]
# 	FormulaDetail.LineType      => 'RawMaterials'[1]
# 	FormulaDetail.Code          => 'RawMaterials'[2]
# 	FormulaDetail.Quantity      => 'RawMaterials'[3]
# 	FormulaDetail.UOM           => 'RawMaterials'[4]
# 	FormulaDetail.UOMDescr      => 'RawMaterials'[5]
# 	FormulaDetail.Comment       => 'RawMaterials'[6]
# 	FormulaDetail.PctWtGross    => 'RawMaterials'[7]
# 	FormulaDetail.Loss          => 'RawMaterials'[8]
# 	Customer.CustomerName       => 'CustomerName'
# 	FormulaObjectives.Objective => 'Objectives'[0]
# 	FormulaObjectives.Target    => 'Objectives'[1]
# 	FormulaObjectives.TargetUnits   => 'Objectives'[2]
# 	FormulaObjectives.LossPercent   => 'Objectives'[3]
# 	FormulaObjectives.InertItemCode => 'Objectives'[4]
# 	FormulaObjectives.Actual    => 'Objectives'[5]

	my $class = shift;
	my @args  = @_;
	my $self  = {};
	
	$self->{'FormulaCode'} = $args[0];
	my $dbh = FormulaTool->connect_formdb();
	my @results = FormulaTool->querydb($dbh,'FormulaCode,Description,Class,SubClass,ServingSize,ServingUOM,Status,Memo01,Memo02,Memo03,AddedBy,Added,DefaultUOM,FormulaTotal,FormulaUOM,CustomerCode,Memo05','FormulaMaster',"FormulaCode like '$self->{\"FormulaCode\"}'");
    
    $results[0][0][1] =~ s/:/-/m;
	$self->{'Description'}	= FormulaTool->despace($results[0][0][1]);
	$self->{'Class'}	= FormulaTool->despace($results[0][0][2]);
	$self->{'SubClass'}	= FormulaTool->despace($results[0][0][3]);
 	$self->{'ServingTotal'}	= sprintf "%0.8f", $results[0][0][4];
 	$self->{'ServingUOM'}	= $results[0][0][5];
 	$self->{'Status'}	= $results[0][0][6];
 	$self->{'ServingDesc'}	= FormulaTool->despace($results[0][0][7]);
 	$self->{'ServingSize'}	= FormulaTool->despace($results[0][0][8]);
 	$self->{'ServingQty'}	= $results[0][0][9];
 	$self->{'Appearance'}	= $results[0][0][16];
 	$self->{'AddedBy'}	= $results[0][0][10];
 	$self->{'Added'}	= $results[0][0][11];
 	$self->{'DefaultUOM'}	= $results[0][0][12];
 	$self->{'FormulaTotal'}	= sprintf "%0.8f", $results[0][0][13];
 	$self->{'FormulaUOM'}	= $results[0][0][14];
 	$self->{'CustomerCode'}	= FormulaTool->despace($results[0][0][15]);
    
	@results = FormulaTool->querydb($dbh,'LineNumber,LineType,Code,Quantity,UOM,UOMDescr,Comment,PctWtGross,Loss','FormulaDetail',"FormulaCode like '$self->{\"FormulaCode\"}' AND LineType = 1");
	$self->{'RawMaterials'}	= \@{ $results[0] };
    
	@results = FormulaTool->querydb($dbh,'CustomerCode, CustomerName','Customer',"CustomerCode like '$self->{\"CustomerCode\"}'");
	$self->{'CustomerName'}	= FormulaTool->despace($results[0][0][1]);
    
	@results = FormulaTool->querydb($dbh,'Objective, Target, TargetUnits, LossPercent, InertItemCode, Actual','FormulaObjectives',"FormulaCode like '$self->{\"FormulaCode\"}'");
	$self->{'Objectives'}	= \@{ $results[0] };

	return bless $self, $class;
}

1;