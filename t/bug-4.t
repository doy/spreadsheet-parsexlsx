#!/usr/bin/env perl
use strict;
use warnings;
use Test::More;

use Spreadsheet::ParseXLSX;

my $parser = Spreadsheet::ParseXLSX->new;
my $wb = $parser->parse('t/data/bug-4.xlsx');
my $ws = $wb->worksheet(0);

{
    my $cell = $ws->get_cell(0, 0);
    is($cell->value,'Order Number', 'cell check 1');
}

{
    my $cell = $ws->get_cell(1, 0);
    is($cell->value,'364968', 'cell check 2');
}

done_testing;
