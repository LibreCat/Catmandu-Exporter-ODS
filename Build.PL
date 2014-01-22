use strict;
use warnings;
use Module::Build;

my $builder = Module::Build->new(
    module_name => 'Catmandu::Exporter::ODS',
    license => 'perl',
    dist_author => [
        'Snorri Briem <snorri.briem@ub.lu.se>'
    ],
    dist_version_from => 'lib/Catmandu/Exporter/ODS.pm',
    build_requires => {
        'Software::License' => 0,
        'Test::Exception' => 0,
        'Test::More' => 0,
    },
    requires => {
        'perl' => '5.10.1',
        'Archive::Zip' => '1.30' ,
        'Catmandu' => '0.0402',
        'Moo' => '1.0',
        'Spreadsheet::Wright' => '0.104',
    },
    add_to_cleanup => [qw(
        Catmandu-Exporter-ODS-*
    )],
    create_makefile_pl => 'traditional',
    create_license => 1,
);

$builder->create_build_script;
