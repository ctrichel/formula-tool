package FormRaw;

use strict;
our $VERSION = '1.3.7';

use FormulaTool;

sub new {
# Formulator Raw Material
#	ItemCode
#	RawMaterials.Description    => 'Description'
#	RawMaterials.Class      => 'Class'
#	RawMaterials.SubClass   => 'SubClass'
#	RawMaterials.Status     => 'Status'
#	RawMaterials.Cost       => 'Cost'
#	RawMaterials.CostUOM    => 'CostUOM'
#	RawMaterials.CostUOMDescr   => 'CostUOMDescr'
#	RawMaterials.IsRaw      => 'IsRaw'
#	RawMaterials.VendorCode => 'VendorCode'
#	RawMaterials.DisplayDensity => 'DisplayDensity'
#	RawMaterials.ServingSize    => 'ServingSize'
#	RawMaterials.ServingUOM => 'ServingUOM'
#	RawMaterials.LotTracked => 'LotTracked'
#	RawMaterials.AddedBy    => 'AddedBy'
#	RawMaterials.Added      => 'Added'
#	RawMaterials.UpdatedBy  => 'UpdatedBy'
#	RawMaterials.Updated    => 'Updated'
#	RawNutr_Val.NutrCode    => 'Nutrients'[0]
#	RawNutr_Val.NutrQty     => 'Nutrients'[1]
#	Nutr_Def.Description    => 'Nutrients'[2]
#	RawAllergens.AllergenNo => 'Allergens'[0] wheat,milk,soy,eggs,treenuts,peanuts,fish,seafood,organic,gluten
#	RawAllergens.Contains   => 'Allergens'[1]

	my $class = shift;
	my @args  = @_;
	my $self  = {};

	$self->{'ItemCode'} = $args[0];
	my $dbh = FormulaTool->connect_formdb();
	my @results = FormulaTool->querydb($dbh,'ItemCode,Description,Class,SubClass,Status,Cost,CostUOM,CostUOMDescr,IsRaw,VendorCode,DisplayDensity,ServingSize,ServingUOM,LotTracked,AddedBy,Added,UpdatedBy,Updated','RawMaterials',"ItemCode like '$self->{\"ItemCode\"}'");
    
	$self->{'Description'} = FormulaTool->despace($results[0][0][1]);
	$self->{'Class'} = $results[0][0][2];
	$self->{'SubClass'} = $results[0][0][3];
	$self->{'Status'} = $results[0][0][4];
	$self->{'Cost'} = $results[0][0][5];
	$self->{'CostUOM'} = $results[0][0][6];
	$self->{'CostUOMDescr'} = $results[0][0][7];
	$self->{'IsRaw'} = $results[0][0][8];
	$self->{'VendorCode'} = $results[0][0][9];
	$self->{'DisplayDensity'} = $results[0][0][10];
	$self->{'ServingSize'} = $results[0][0][11];
	$self->{'ServingUOM'} = $results[0][0][12];
	$self->{'LotTracked'} = $results[0][0][13];
	$self->{'AddedBy'} = $results[0][0][14];
	$self->{'Added'} = $results[0][0][15];
	$self->{'UpdatedBy'} = $results[0][0][16];
	$self->{'Updated'} = $results[0][0][17];
    
	@results = FormulaTool->querydb($dbh,'NutrCode,NutrQty','RawNutr_Val',"ItemCode like '$self->{\"ItemCode\"}'");
	foreach my $line ( 0..$#{ $results[0] } ) {
		my @desc = FormulaTool->querydb($dbh,'Description','Nutr_Def',"NutrCode like '$results[0][$line][0]'");
		push @{ $results[0] }[$line],$desc[0][0][0];
	}
	$self->{'Nutrients'} = \@{ $results[0] };
    
	@results = FormulaTool->querydb($dbh,'Collateral','RawAllergens',"ItemCode like '$self->{\"ItemCode\"}' order by AllergenNo");
    $self->{'Allergens'} = \@{ $results[0] };
    my $check = 0;
    for my $count (0 .. $#{ $results[0] }) {
        if ( $results[0][$check][0] ) { $check++; }
    }
    $self->{'AllergenCheck'} = $check;
	return bless $self, $class;
}

1;
 	