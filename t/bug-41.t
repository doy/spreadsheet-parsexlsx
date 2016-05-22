#!/usr/bin/env perl
use strict;
use warnings;
use Test::More;

use Spreadsheet::ParseXLSX;

{
    local $SIG{__WARN__} = sub { fail("unexpected warning: $_[0]") };
    my $wb = Spreadsheet::ParseXLSX->new->parse('t/data/bug-41.xlsx');
    pass('it parses successfully');
}

done_testing;
