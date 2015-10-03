#!/usr/bin/env perl
use strict;
use warnings;
use Test::More;

use Spreadsheet::ParseXLSX;

my $wb;
eval {
    $wb = Spreadsheet::ParseXLSX->new->parse('t/data/target-abspath.xlsx');
};
if ($@) {
    diag $@;
}
ok((not $@), "parsing target-abspath.xlsx ok");

done_testing;
