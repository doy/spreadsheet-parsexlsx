#!/usr/bin/env perl
use strict;
use warnings;
use Test::More;

use Spreadsheet::ParseXLSX;

my $wb = Spreadsheet::ParseXLSX->new->parse('t/data/bug-6.xlsx');
is($wb->worksheet_count, 8);

my %cells = (
    7 => {
        0 => 'mfg fdproc',
        1 => 'Tom Forsythe',
    },
    8 => {
        0 => 'ent bartend-402 data max prodigy max 203 dpi',
        1 => 'Dave Levos ; Tommy Holland',
    },
    9 => {
        0 => 'ent bartend-402 inter px4i 400 dpi rw',
        1 => 'Tommy Holland; Dave Levos',
    },
    10 => {
        0 => 'opr-mfg asmb inst ro',
        1 => 'Chris McGee',
    },
);

my $ws = $wb->worksheet('DSGroups');
for my $row (sort { $a <=> $b } keys %cells) {
    for my $col (sort { $a <=> $b } keys %{ $cells{$row} }) {
        my $cell = $ws->get_cell($row, $col);
        next unless $cell;
        is($cell->value, $cells{$row}{$col}, "correct value for ($row, $col)");
    }
}

done_testing;
