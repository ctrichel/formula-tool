package PProRaw;

use strict;
our $VERSION = '1.3.3';

sub new {
# PPro Raw Material
#	*ItemCode
#	icitem.itmdesc  => 'Description'
#	icitem.stkumid  => 'StockUOM'
#	icitem.sunmsid  => 'SalesUOM'
#	icitem.punmsid  => 'PurchUOM'
#	icitem.cunmsid  => 'CountUOM'
#	icitem.cmpumid  => 'CompUOM'
#	icitem.plinid   => 'ProductLine'
#	icitem.ionhand  => 'QtyOnHand'
#	icitem.decnum   => 'Decimals'
#	icitem.stkcode  => 'Stock?'
#	icitem.type     => 'Type'
#	icitem.resell   => 'Resell?'
#	icitem.uselots  => 'UseLots?'
#	icitem.useserl  => 'UseSerial?'
#	icitem.makeitr  => 'MkInternal?'
#	icitem.itmclss  => 'Class'
#	icitem.code     => 'SubClass'
#	icitem.itemstat => 'Status'
#	icitem.nomonths => 'Expires'
#	icitem.mthsordays   => 'ExpiresUOM'
#	icitem.adduser  => 'AddedBy'
#	icitem.adddate  => 'Added'
#	icitem.purextr  => 'PurExternal?'
#	icitem.lckuser  => 'UpdatedBy'
#	icitem.lckdate  => 'Updated'
#	icitem.avgcost  => 'AvgCost'
#	icitem.stdcost  => 'StdCost'
#	icitem.suplier  => 'VendorCode'
#	iciloc.item     => 'WH1?'
#   iciloc.icacct   => 'GLAccounts'[0]
#   iciloc.rclacct   => 'GLAccounts'[1]
#   iciloc.iclacct   => 'GLAccounts'[2]
#   iciloc.dlbacct   => 'GLAccounts'[3]
#   iciloc.fxoacct   => 'GLAccounts'[4]
#   iciloc.vroacct   => 'GLAccounts'[5]
#   iciloc.mipacct   => 'GLAccounts'[6]
#   iciloc.dipacct   => 'GLAccounts'[7]
#   iciloc.fipacct   => 'GLAccounts'[8]
#   iciloc.vipacct   => 'GLAccounts'[9]
#   iciloc.mclacct   => 'GLAccounts'[10]
#   iciloc.icshpclear   => 'GLAccounts'[11]
#	ppalgn.wheat    => 'Allergens'[0]
#	ppalgn.milk     => 'Allergens'[1]
#	ppalgn.soy      => 'Allergens'[2]
#	ppalgn.eggs     => 'Allergens'[3]
#	ppalgn.treenuts => 'Allergens'[4]
#	ppalgn.peanuts  => 'Allergens'[5]
#	ppalgn.fish     => 'Allergens'[6]
#	ppalgn.seafood  => 'Allergens'[7]
#	ppalgn.algn1	=> 'Allergens'[8]
#	ppalgn.algn2	=> 'Allergens'[9]

	my $class = shift;
	my @args  = @_;
	my $self  = {};
	
	$self->{'ItemCode'} = $args[0];
	my $dbh = FormulaTool->connect_pprodb();
	my @results = FormulaTool->querydb($dbh,'item,itmdesc,stkumid,sunmsid,punmsid,cunmsid,cmpumid,plinid,ionhand,decnum,stkcode,type,resell,uselots,useserl,makeitr,itmclss,code,itemstat,nomonths,mthsordays,adduser,adddate,purextr,lckuser,lckdate,lstcost,stdcost,suplier','icitem',"item like '$self->{\"ItemCode\"}'");
    
	$self->{'Description'}	= FormulaTool->despace($results[0][0][1]);
	$self->{'StockUOM'}	= $results[0][0][2];
	$self->{'SalesUOM'}	= $results[0][0][3];
	$self->{'PurchUOM'}	= $results[0][0][4];
	$self->{'CountUOM'}	= $results[0][0][5];
	$self->{'CompUOM'}	= $results[0][0][6];
	$self->{'ProductLine'}	= $results[0][0][7];
	$self->{'QtyOnHand'}	= $results[0][0][8];
	$self->{'Decimals'}	= $results[0][0][9];
	$self->{'Stock?'}	= $results[0][0][10];
	$self->{'Type'}		= $results[0][0][11];
	$self->{'Resell?'}	= $results[0][0][12];
	$self->{'UseLots?'}	= $results[0][0][13];
	$self->{'UseSerial?'}	= $results[0][0][14];
	$self->{'MkInternal?'}	= $results[0][0][15];
	$self->{'Class'}	= $results[0][0][16];
    $self->{'SubClass'} = $results[0][0][17];
	$self->{'Status'}	= $results[0][0][18];
	$self->{'Expires'}	= $results[0][0][19];
	$self->{'ExpiresUOM'}	= $results[0][0][20];
	$self->{'AddedBy'}	= $results[0][0][21];
	$self->{'Added'}	= $results[0][0][22];
	$self->{'PurExternal?'}	= $results[0][0][23];
	$self->{'UpdatedBy'}	= $results[0][0][24];
	$self->{'Updated'}	= $results[0][0][25];
	$self->{'AvgCost'}	= $results[0][0][26];
	$self->{'StdCost'}	= $results[0][0][27];
	$self->{'VendorCode'}	= $results[0][0][28];
    
    @results = FormulaTool->querydb( $dbh,'icacct,rclacct,iclacct,dlbacct,fxoacct,vroacct,mipacct,dipacct,fipacct,vipacct,mclacct,icshpclear','iciloc',"item like '$self->{\"ItemCode\"}'");
    $self->{'WH1?'} = ($results[0][0][0]?1:0);
    $self->{'GLAccounts'} = \@{ $results[0][0] };
    
	@results = FormulaTool->querydb( $dbh,'wheat,milk,soy,eggs,treenuts,peanuts,fish,seafood','PPALGN',"item like '$self->{\"ItemCode\"}'" );
	$self->{'Allergens'}	= \@{ $results[0][0] };
    my $check = 0;
    for my $count (0 .. $#{ $results[0][0] }) {
        if ( $results[0][0][$check] > 0) { $check++; }
    }
    $self->{'AllergenCheck'} = $check;
    
	return bless $self, $class;
}

1;