package Batch;

use strict;
our $VERSION = '1.1.3';

use FormulaTool;
    
sub new {
# Batch (WorkOrder)
# 	BatchRecord
#   PriorLot
#   Yield
# 	ppwohd.item     => 'FormulaCode'
# 	ppwohd.lot      => 'Lot'
# 	ppwohd.status   => 'Status'
# 	ppwohd.wostat   => 'WOStatus'
# 	ppwohd.qty      => 'AllocQty'
# 	ppwohd.projqty  => 'NeedQty'
# 	ppwohd.route    => 'Route'
# 	ppwohd.reqdte   => 'ReqDate'
# 	ppwohd.version  => 'Version'
# 	ppwohd.compqty  => 'CompQty'
#   ppwohd.clsdate  => 'CloseDate'
# 	ppwodt.trancode => 'BatchDetails'[0]
# 	ppwodt.trandte  => 'BatchDetails'[1]
# 	ppwodt.fitem    => 'BatchDetails'[2]
# 	ppwodt.flotno   => 'BatchDetails'[3]
# 	ppwodt.partno   => 'BatchDetails'[4]
# 	ppwodt.qty      => 'BatchDetails'[5]
# 	ppwodt.qtyaloc  => 'BatchDetails'[6]

	my $class = shift;
	my @args  = @_;
	my $self  = {};
	
	$self->{'BatchRecord'} = $args[0];
	my $dbh = FormulaTool->connect_pprodb();
    my @results = FormulaTool->querydb($dbh,'wono,item,lot,status,wostat,qty,projqty,route,reqdte,version,compqty,clsdate','ppwohd',"wono like '%$self->{\"BatchRecord\"}'");
    
	$self->{'FormulaCode'} = FormulaTool->despace( $results[0][0][1] );
	if (int($results[0][0][2]) > 0) {
        $self->{'Lot'} = int($results[0][0][2]);
        my @tresult = FormulaTool->querydb($dbh,'LOTNO', 'ICILOT',"(ITEM LIKE '$self->{\"FormulaCode\"}') and (LOTNO < $self->{\"Lot\"}) order by EXPIRES desc" );
        $self->{'PriorLot'} = $tresult[0][0][0];
    } else {
        $self->{'Lot'} = "-";
        $self->{'PriorLot'} = "-";
    }
    $self->{'Status'} = $results[0][0][3];
	$self->{'WOStatus'} = $results[0][0][4];
	$self->{'AllocQty'} = sprintf "%0.4f",$results[0][0][5];
	$self->{'NeedQty'} = sprintf "%0.4f",$results[0][0][6];
	$self->{'Route'} = $results[0][0][7];
    $self->{'RouteSize'} = int ((split q{ }, $results[0][0][7])[1]);
    if (!$self->{'RouteSize'}) { $self->{'RouteSize'} = 1; }
	$self->{'ReqDate'} = $results[0][0][8];
	$self->{'Version'} = $results[0][0][9];
	$self->{'CompQty'} = sprintf "%0.4f",$results[0][0][10];
    $self->{'CloseDate'} = $results[0][0][11];
    
	@results = FormulaTool->querydb($dbh,'trancode,trandte,fitem,flotno,partno,qty,qtyaloc','ppwodt',"wono like '%$self->{\"BatchRecord\"}'");
	$results[0][0][4] = FormulaTool->despace($results[0][0][4]);    
	$self->{'BatchDetails'}	= \@{ $results[0] };
    
    @results = FormulaTool->querydb($dbh,'yield','ppbmhd',"item like '$self->{\"FormulaCode\"}' and version like '%$self->{\"Version\"}'");
    $self->{'Yield'} = $results[0][0][0] / 100;
    
	return bless $self, $class;
}

1;