package SpecFormula;

use strict;
our $VERSION = '1.0.0';

use FormulaTool;

sub new {
# Specification Formula
#	FormulaCode
#	FP Specs.Customer Name     => 'Customer'
#	FP Specs.Product Name      => 'ProductName'
#	FP Specs.Product Type      => 'ProductType'
#	FP Specs.UnitType          => 'UnitTypeSpec'
#	FP Specs.UnitSizeShape     => 'UnitSizeShapeSpec'
#	FP Specs.UnitWeightSpec    => 'UnitWeightSpec'
#	FP Specs.UnitWeightRange   => 'UnitWeightRangeSpec'
#	FP Specs.UnitHardnessSpec  => 'UnitHardnessSpec'
#	FP Specs.UnitThicknessSpec => 'UnitThicknessSpec'
#	FP Specs.UnitLengthSpec    => 'UnitLengthSpec'
#	FP Specs.UnitDisintegrationSpec=> 'UnitDisintegrationSpec'
#	FP Specs.UnitCoatingSpec   => 'UnitCoatingSpec'
#	FP Specs.UnitAppearance    => 'UnitAppearanceSpec'
#	FP Specs.UnitPerServing    => 'UnitsPerSvg'
#	FP Specs.MaxUnitsPerDay    => 'MaxUnitsPerDay'
#	FP Specs.UnitpH            => 'UnitpHSpec'
#	FP Specs.UnitTaste         => 'UnitTasteSpec'
#	FP Specs.UnitVegetarian    => 'UnitVegetarianSpec'
#	FP Specs.UnitVegan         => 'UnitVeganSpec'
#	FP Specs.UnitKosher        => 'UnitKosherSpec'
#	FP Specs.UnitHalal         => 'UnitHalalSpec'
#	FP Specs.UnitGluten        => 'UnitGlutenSpec'
#	FP Specs.UnitNonGMO        => 'UnitNonGMOSpec'
#	FP Specs.UnitReqSpec       => 'UnitSpecReqs'
#	FP Specs.UnitCertOrganic   => 'UnitOrganicSpec'


	my $class = shift;
	my @args  = @_;
	my $self  = {};

	$self->{'FormulaCode'} = $args[0];
	my $dbh = FormulaTool->connect_specdb();
    #|         ID         |     Formula ID     |   Customer Name    |    Product Name    |    Product Type    |      UnitType      |   UnitSizeShape    |   UnitWeightSpec   |  UnitWeightRange   |  UnitHardnessSpec  | UnitThicknessSpec  |   UnitLengthSpec   | UnitDisintegrationSpec |  UnitCoatingSpec   |   UnitAppearance   |  UnitsPerServing   |   MaxUnitsPerDay   |     Active Sku     |       UnitpH       |     UnitTaste      |   UnitVegetarian   |     UnitVegan      |     UnitKosher     |     UnitHalal      |     UnitGluten     |     UnitNonGMO     |    UnitReqSpec     |  UnitCertOrganic   |
	my @results = FormulaTool->querydb($dbh,'[Customer Name],[Product Name],[Product Type],UnitType,UnitSizeShape,UnitWeightSpec,UnitWeightRange,UnitHardnessSpec,UnitThicknessSpec,UnitLengthSpec,UnitDisintegrationSpec,UnitCoatingSpec,UnitAppearance,UnitsPerServing,MaxUnitsPerDay,UnitpH,UnitTaste,UnitVegetarian,UnitVegan,UnitKosher,UnitHalal,UnitGluten,UnitNonGMO,UnitReqSpec,UnitCertOrganic','[FP Specs]',"[Formula ID] like '$self->{\"FormulaCode\"}'");
    
	$self->{'Customer'} = FormulaTool->despace($results[0][0][0]);
	$self->{'ProductName'} = FormulaTool->despace($results[0][0][1]);
	$self->{'ProductType'} = $results[0][0][2];
	$self->{'UnitTypeSpec'} = $results[0][0][3];
	$self->{'UnitSizeShapeSpec'} = $results[0][0][4];    
	$self->{'UnitWeightSpec'} = $results[0][0][5];
	$self->{'UnitWeightRangeSpec'} = $results[0][0][6];    
	$self->{'UnitHardnessSpec'} = $results[0][0][7];
	$self->{'UnitThicknessSpec'} = $results[0][0][8];    
	$self->{'UnitLengthSpec'} = $results[0][0][9];
	$self->{'UnitDisintegrationSpec'} = $results[0][0][10];    
	$self->{'UnitCoatingSpec'} = $results[0][0][11];
	$self->{'UnitAppearanceSpec'} = $results[0][0][12];    
	$self->{'UnitsPerSvg'} = $results[0][0][13];
	$self->{'MaxUnitsPerDay'} = $results[0][0][14];    
	$self->{'UnitphSpec'} = $results[0][0][15];
	$self->{'UnitTasteSpec'} = $results[0][0][16];    
	$self->{'UnitVegetarianSpec'} = $results[0][0][17];
	$self->{'UnitVeganSpec'} = $results[0][0][18];    
	$self->{'UnitKosherSpec'} = $results[0][0][19];
	$self->{'UnitHalalSpec'} = $results[0][0][20];    
	$self->{'UnitGlutenSpec'} = $results[0][0][21];
	$self->{'UnitNonGMOSpec'} = $results[0][0][22];    
	$self->{'UnitSpecReqs'} = $results[0][0][23];
	$self->{'UnitOrganicSpec'} = $results[0][0][24]; 
    
	return bless $self, $class;
}

1;
 	