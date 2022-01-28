package PProFormula;

use strict;
our $VERSION = '1.3.0';

use FormulaTool;
use FormulaTool::PProRaw;

sub new {
# PPro Formula
# 	FormulaCode
# 	ppbmhd.version      => 'Version'
# 	PProPRWeight        => 'PRWeight'
# 	ppbmhd.revstat      => 'RevStatus'
# 	ppbmhd.revdate      => 'RevDate'
# 	ppbmhd.inactive     => 'Inactive?'
# 	ppbmhd.type         => 'Type'
# 	ppbmhd.route        => 'Route'
# 	ppbmhd.partial      => 'Partial?'
# 	ppbmhd.batches      => 'Batches'
# 	ppbmhd.yield        => 'Yield'
# 	ppbmhd.overhead     => 'Overhead'
# 	ppbmhd.revision     => 'Revision'
# 	ppbmhd.lckdate      => 'Updated'
# 	ppbmhd.lckuser      => 'UpdatedBy'
# 	ppbmhd.adddate      => 'Added'
# 	ppbmhd.adduser      => 'AddedBy'
# 	ppbmdt.partno       => 'Details'[0]
# 	ppbmdt.findnum      => 'Details'[1]
# 	ppbmdt.qty          => 'Details'[2]
# 	ppbmdt.adduser      => 'Details'[3]
# 	ppbmdt.adddate      => 'Details'[4]
# 	icilot.lotno        => 'Lots'[0]
# 	icilot.odocno       => 'Lots'[1]
# 	icilot.mfgdate      => 'Lots'[2]
#   ppwohd.wono         => 'Wono'[0]
#   ppwohd.lot          => 'Wono'[1]
#   ppwohd.clsdate      => 'Wono'[2]

	my $class = shift;
	my @args  = @_;
	my $self  = {};
	
	$self->{'FormulaCode'} = $args[0];
	my $dbh = FormulaTool->connect_pprodb();
    my @results = FormulaTool->querydb($dbh,'itmdesc','icitem',"item like '$self->{\"FormulaCode\"}'");
    $results[0][0][0] =~ s/:/-/m;
    
    $self->{'Description'} = FormulaTool->despace($results[0][0][0]);
	@results = FormulaTool->querydb($dbh,'version','ppbmhd',"item like '$self->{\"FormulaCode\"}'");
    $self->{'Version'} = '01';
    if ($#{ $results[0] } > 0) {
        print "There are ".($#{ $results[0] } +1)." versions of this formula in PPro.\nVersions:";
        foreach my $count (0..$#{ $results[0] }) {
            printf "\t%02d",$results[0][$count][0];
        }
        while(1) {
            print "\n\n\tEnter the Version Number to use [default: $results[0][0][0]] : ";
            my $input = <>;
            chomp $input;
            print "\n";
        
            if ( !$input ) {
                $self->{'Version'} = sprintf "%02d",$results[0][0][0];
                last;
            }
            elsif ( ( $input !~ m/^\d+$/i ) || ( $input > 9 ) ) {
                print "\n$input is an invalid version.\n";
                next;
            }
            else {
                $self->{'Version'} = sprintf "%02d",$input;
                last;
            }
        }
	}

	@results = FormulaTool->querydb($dbh,'item,revstat,revdate,inactive,type,route,partial,batches,yield,overhead,revision,lckdate,lckuser,adddate,adduser','ppbmhd',"item like '$self->{\"FormulaCode\"}' AND version like '$self->{'Version'}'");
	$self->{'RevStatus'}	= $results[0][0][1];
    if ($results[0][0][2]) { $self->{'RevDate'}	= $results[0][0][2]; } else { $self->{'RevDate'}	= '1900-01-01'; }
    $self->{'Inactive?'}	= $results[0][0][3];
	$self->{'Type'}		= $results[0][0][4];
	$self->{'Route'}	= FormulaTool->despace($results[0][0][5]);
	$self->{'Partial?'}	= $results[0][0][6];
	$self->{'Batches'}	= $results[0][0][7];
	$self->{'Yield'}	= $results[0][0][8];
	$self->{'Overhead'}	= FormulaTool->despace($results[0][0][9]);
    $self->{'Revision'} = ((defined $results[0][0][11])?FormulaTool->despace($results[0][0][10]):"0");
	$self->{'Updated'}	= $results[0][0][11];
	$self->{'UpdatedBy'}	= $results[0][0][12];
	$self->{'Added'}	= $results[0][0][13];
	$self->{'AddedBy'}	= $results[0][0][14];
    
	@results = FormulaTool->querydb($dbh,'lstcost','icitem',"item like '$self->{\"FormulaCode\"}'");
	$self->{'LastCost'}	= $results[0][0][0];    
    
	@results = FormulaTool->querydb($dbh,'partno,findnum,qty,adduser,adddate','ppbmdt',"item like '$self->{\"FormulaCode\"}' AND version like '$self->{'Version'}'");
	$self->{'Details'}	= \@{ $results[0] };
    my %allergencount;
    foreach my $item ( 0 .. $#{ $results[0] } ) {
        my $item_obj = PProRaw->new( FormulaTool->despace($self->{'Details'}[$item][0]) );
        if ( $item_obj->{'Allergens'}[0] ) {
            $self->{'Allergens'}{ FormulaTool->despace($self->{'Details'}[$item][0]) } = 'Wheat';
            ++$allergencount{'Wheat'};
        }
        if ( $item_obj->{'Allergens'}[1] ) {
            $self->{'Allergens'}{ FormulaTool->despace($self->{'Details'}[$item][0]) } = 'Milk';
            ++$allergencount{'Milk'};
        }
        if ( $item_obj->{'Allergens'}[2] ) {
            $self->{'Allergens'}{ FormulaTool->despace($self->{'Details'}[$item][0]) } = 'Soy';
            ++$allergencount{'Soy'};
        }
        if ( $item_obj->{'Allergens'}[3] ) {
            $self->{'Allergens'}{ FormulaTool->despace($self->{'Details'}[$item][0]) } = 'Eggs';
            ++$allergencount{'Eggs'};
        }
        if ( $item_obj->{'Allergens'}[4] ) {
            $self->{'Allergens'}{ FormulaTool->despace($self->{'Details'}[$item][0]) } = 'Treenuts';
            ++$allergencount{'Treenuts'};
        }
        if ( $item_obj->{'Allergens'}[5] ) {
            $self->{'Allergens'}{ FormulaTool->despace($self->{'Details'}[$item][0]) } = 'Peanuts';
           ++$allergencount{'Peanuts'};
        }
        if ( $item_obj->{'Allergens'}[6] ) {
            $self->{'Allergens'}{ FormulaTool->despace($self->{'Details'}[$item][0]) } = 'Fish';
            ++$allergencount{'Fish'};
        }
        if ( $item_obj->{'Allergens'}[7] ) {
            $self->{'Allergens'}{ FormulaTool->despace($self->{'Details'}[$item][0]) } = 'Shellfish';
            ++$allergencount{'Shellfish'};
        }
    }    
    $self->{'AllergenCount'} = \%allergencount;
    foreach my $type ( keys $self->{'AllergenCount'} ) {
        if ( $self->{'AllergenCount'}{$type} > 0 ) { $self->{'AllergenList'} .= "$type, " }
    }
    if ( !$self->{'AllergenList'} ) { $self->{'AllergenList'} = 'None'; }

	@results = FormulaTool->querydb($dbh,'lotno,odocno,mfgdate','icilot',"item like '$self->{\"FormulaCode\"}' order by cast(LOTNO as INT) desc, MFGDATE desc");
	$self->{'Lots'}		= \@{ $results[0] };
    
	@results = FormulaTool->querydb($dbh,'item,wono,lot,clsdate,wostat','ppwohd',"item like '$self->{\"FormulaCode\"}' and wostat < 5 order by cast(WONO as INT) desc, CLSDATE desc");
	$self->{'Wono'}		= \@{ $results[0] };
    
	@results = FormulaTool->querydb($dbh,'item,wono,lot,clsdate,wostat','ppwohd',"item like '$self->{\"FormulaCode\"}\.PR' and wostat < 5 order by cast(WONO as INT) desc, CLSDATE desc");
	$self->{'PRWono'}		= \@{ $results[0] };
    
    @results = FormulaTool->querydb($dbh,'item,wono,lot,clsdate,wostat','ppwohd',"(item like '$self->{\"FormulaCode\"}\%' AND item NOT LIKE '$self->{\"FormulaCode\"}\.PR' AND item NOT LIKE '$self->{\"FormulaCode\"}') and wostat < 5 order by cast(WONO as INT) desc, CLSDATE desc");
	$self->{'PKGWono'}		= \@{ $results[0] };
    
    @results = FormulaTool->querydb($dbh,'version','ppbmhd',"item like '$self->{\"FormulaCode\"}\.PR'");
    my $prversion = '01';
    if ($#{ $results[0] } > 0) {
        print "There are ".($#{ $results[0] } +1)." versions of this PR in PPro.\nVersions:";
        foreach my $count (0..$#{ $results[0] }) {
            printf "\t%02d",$results[0][$count][0];
        }
        while(1) {
            print "\n\n\tEnter the Version Number to use [default: $results[0][0][0]] : ";
            my $input = <>;
            chomp $input;
            print "\n";
        
            if ( !$input ) {
                $prversion = sprintf "%02d",$results[0][0][0];
                last;
            }
            elsif ( ( $input !~ m/^\d+$/i ) || ( $input > 9 ) ) {
                print "\n$input is an invalid PR version.\n";
                next;
            }
            else {
                $prversion = sprintf "%02d",$input;
                last;
            }
        }
	}

    @results = FormulaTool->querydb($dbh,'QTY','PPBMDT',"ITEM like '$self->{\"FormulaCode\"}.PR' AND PARTNO like '$self->{\"FormulaCode\"}' AND VERSION = '$prversion'");
    if ($results[0][0][0]) {
        $self->{'PRWeight'} = sprintf "%0.8f", $results[0][0][0];
    }

	return bless $self, $class;
}

1;

 	