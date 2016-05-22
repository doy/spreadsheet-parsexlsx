#!/usr/bin/env perl
use strict;
use warnings;
use Test::More;

use Spreadsheet::ParseXLSX;

{
    my $wb = Spreadsheet::ParseXLSX->new->parse('t/data/bug-32.xlsx');

    my $ws1 = $wb->worksheet(0);
    like($ws1->get_cell(0, 0)->value, qr/^PURSUANT/);

    my $ws2 = $wb->worksheet(1);
    like($ws2->get_cell(0, 0)->value, qr/^QMS/);
}

{
    my $wb = Spreadsheet::ParseXLSX->new->parse('t/data/bug-32-2.xlsx');

    my $ws = $wb->worksheet(0);
    is($ws->get_cell(1, 1)->value, 93);
}

done_testing;
