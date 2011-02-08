use strict;
use warnings;
use ExtUtils::MakeMaker;

WriteMakefile(
    NAME                => 'Tidy::Spreadsheet',
    AUTHOR              => q{Ashley Pope <ashleyp@cpan.org>},
    VERSION_FROM        => 'lib/Tidy/Spreadsheet.pm',
    ABSTRACT_FROM       => 'lib/Tidy/Spreadsheet.pm',
    ($ExtUtils::MakeMaker::VERSION >= 6.3002
      ? ('LICENSE'=> 'perl')
      : ()),
    PL_FILES            => {},
    PREREQ_PM => {
        'Test::More' => 0,
    },
    dist                => { COMPRESS => 'gzip -9f', SUFFIX => 'gz', },
    clean               => { FILES => 'Tidy-Spreadsheet-*' },
);
