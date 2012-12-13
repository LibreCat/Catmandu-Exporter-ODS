package Catmandu::Exporter::ODS;

use Catmandu::Sane;
use Moo;
use Spreadsheet::Wright;
use Archive::Zip;

with 'Catmandu::Exporter';

our $VERSION = '0.02';

has ods       => ( is => 'ro', lazy => 1, builder => '_build_ods' );
has header    => ( is => 'ro', default => sub { 1 } );
has temp_file => ( is => 'ro', lazy => 1, builder => '_build_temp_file' );
has temp_bufflen => ( is => 'rw', default => sub { 65536 } );

has fields => (
    is     => 'rw',
    coerce => sub {
        my $fields = $_[0];
        given ( ref $fields ) {
            when ('ARRAY') { return $fields }
            when ('HASH') { return [ keys %$fields ] }
            default { return [ split ',', $fields ] }
        }
    },
);

sub _build_temp_file {
    return unless ref $_[0]->file;
    ( Archive::Zip::tempFile() )[1];
}

sub _build_ods {
    my $ods = Spreadsheet::Wright->new(
        file => $_[0]->temp_file || $_[0]->file,
        format => 'ods',
    );
}

sub encoding { ':raw' }

sub add {
    my ( $self, $data ) = @_;
    my $header = $self->header;
    my $fields = $self->fields || $self->fields($data);
    my $ods    = $self->ods;
    my $n      = $self->count;
    if ( $header && $n == 0 ) {
        my $header_labels =
          ref $header
          ? [ map { $header->{$_} // $_ } @$fields ]
          : $fields;
        $ods->addrow(@$header_labels);
    }
    $n++;
    $ods->addrow( @{$data}{@$fields} );
}

sub commit {
    my $self = shift;
    $self->ods->close;

    my $temp_file = $self->temp_file;
    if ($temp_file) {
        my $temp_fh;
        open( $temp_fh, "<:raw", $temp_file )
          or die("Failed to open temporary file '$temp_file': $!");
        my $out_fh = $self->fh;
        my $buffer;
        my $bufflen = $self->temp_bufflen;
        print( $out_fh $buffer ) while ( read( $temp_fh, $buffer, $bufflen ) );
        close $temp_fh;
        unlink $self->temp_file;
    }
}

=head1 NAME

Catmandu::Exporter::ODS - an ODS exporter

=head1 SYNOPSIS

    use Catmandu::Exporter::ODS;

    my $exporter = Catmandu::Exporter::ODS->new(
				file => 'output.ods',
				fix => 'myfix.txt'
				header => 1);

    $exporter->fields("f1,f2,f3");

    $exporter->add_many($arrayref);
    $exporter->add_many($iterator);
    $exporter->add_many(sub { });

    $exporter->add($hashref);

    $exporter->commit;

    printf "exported %d objects\n" , $exporter->count;

=head1 METHODS

=head2 new(header => 0|1|HASH, fields => ARRAY|HASH|STRING)

Creates a new Catmandu::Exporter::ODS. A header line with field names will be
included if C<header> is set. Field names can be read from the first item
exported or set by the fields argument (see: C<fields>).

=head2 fields($arrayref)

Set the field names by an ARRAY reference.

=head2 fields($hashref)

Set the field names by the keys of a HASH reference.

=head2 fields($string)

Set the fields by a comma delimited string.

=head2 header(1)

Include a header line with the field names

=head2 header($hashref)

Include a header line with custom field names

=head2 commit

Commit the changes and close the ODS.

=head1 SEE ALSO

L<Catmandu::Exporter>

=cut

1;
