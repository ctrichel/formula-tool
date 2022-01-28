package LogFile;

use strict;
our $VERSION = '1.1.1';

sub new {
# Log File
# 	Path
# 	*Name
#	FilePath

	my $class = shift;
	my @args = @_;
	my $self = {};
	
	$self->{'Name'} = $args[0];
	if (!$args[0]) { $self->{'Name'} = "logfile.log"; }
	$self->{'Path'} = $args[1];
	$self->{'FilePath'} = $self->{'Path'}.$self->{'Name'};
    
    open (STDERR, ">>", $self->{'Path'}."errors.log") || die "can't redirect STDERR";
    STDERR->autoflush(1);

	return bless $self, $class;
}

sub append {
#	Displays text and appends text to file without timestamp
#	Optional, flag indicates whether to exit app (0) or not (1) or no echo (2)
#
    my $self = shift;
    my @args = @_;
    my $outtext = $_[0];
    my $flag    = $_[1];
    
    open my $FH, '>>', $self->{'FilePath'} or die "!!!unable to append to $self->{'FilePath'}\n$!\n";
    print {$FH} "$outtext\n";
    if ( !$flag ) { die "$outtext\n"; }
    if ( $flag < 2 ) { print "$outtext\n"; }
    close $FH;
    return 1;
}

sub write {
#	Displays text and prints text to new file
#	Optional, flag indicates whether to exit app (0) or not (1) or no echo (2)
#
    my $self = shift;
    my @args = @_;
    my $outtext = $_[0];
    my $flag    = $_[1];
    
    open my $FH, '>', $self->{'FilePath'} or die "!!!unable to write to $self->{'FilePath'}\n$!\n";
    print {$FH} "$outtext\n";
    if ( !$flag ) { die "$outtext\n"; }
    if ( $flag < 2 ) { print "$outtext\n"; }
    close $FH;
    return 1;
}

sub log {
#	Displays text and appends text to log with timestamp and username
#	Optional, flag indicates whether to exit app (0) or not (1) or no echo (2)
#
    my $self = shift;
    my @args = @_;
    my $outtext = $_[0];
    my $flag    = $_[1];
    my $username  = uc $ENV{USERNAME};
    
    open my $FH, '>>', $self->{'FilePath'} or die "!!!unable to log to $self->{'FilePath'}\n$!\n";
    print {$FH} localtime . " $username $outtext\n";
    if ( !$flag ) { die "$outtext\n"; }
    if ( $flag < 2 ) { print "$outtext\n"; }
    close $FH;
    return 1;
}

1;

 	