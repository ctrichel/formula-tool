package FormPkg;

use strict;
our $VERSION = '1.0.2';

use FormulaTool;

sub new {
# Formulator Packaging
#	ItemCode
#	BOMPackaging.InstrText  => 'Description'
#	BOMPackaging.Description    => 'CostUOMDescr'
#	BOMPackaging.Cost       => 'Cost'
#	BOMPackaging.Status     => 'Status'
#	BOMPackaging.PurchasingFactor    => 'PurchFactor'
#	BOMPackaging.VendorCode => 'VendorCode'
#	BOMPackaging.AddedBy    => 'AddedBy'
#	BOMPackaging.Added      => 'Added'
#	BOMPackaging.UpdatedBy  => 'UpdatedBy'
#	BOMPackaging.Updated    => 'Updated'
#	BOMPackaging.Class      => 'Class'
#	BOMPackaging.SubClass   => 'SubClass'

	my $class = shift;
	my @args  = @_;
	my $self  = {};
	
	$self->{'ItemCode'} = $args[0];
	my $dbh = FormulaTool->connect_formdb();
	my @results = FormulaTool->querydb($dbh,'InstrCode,InstrText,Description,Cost,Status,PurchasingFactor,VendorCode,AddedBy,Added,UpdatedBy,Updated,Class,SubClass','BOMPackaging',"InstrCode like '$self->{\"ItemCode\"}'");

    $self->{'Description'} = FormulaTool->despace($results[0][0][1]);
	$self->{'CostUOMDescr'} = $results[0][0][2];
	$self->{'Cost'} = $results[0][0][3];
	$self->{'Status'} = $results[0][0][4];
	$self->{'PurchFactor'} = $results[0][0][5];
	$self->{'VendorCode'} = $results[0][0][6];
	$self->{'AddedBy'} = $results[0][0][7];
	$self->{'Added'} = $results[0][0][8];
	$self->{'UpdatedBy'} = $results[0][0][9];
	$self->{'Updated'} = $results[0][0][10];
	$self->{'Class'} = $results[0][0][11]; 
	$self->{'SubClass'} = $results[0][0][12];
    
	return bless $self, $class;
}

1;