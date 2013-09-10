#!/usr/bin/env perl
use strict;
use warnings;
use Test::More;

use Spreadsheet::ParseXLSX;

my $wb = Spreadsheet::ParseXLSX->new->parse('t/data/bug-6-2.xlsx');
is($wb->worksheet_count, 9);

my %cells = (
    7 => {
        0 => "acr business objects users",
        1 => "MBX or Distribution group",
    },
    8 => {
        0 => "atst",
        1 => "Kevin Krause; Gale Wilson",
    },
    9 => {
        0 => "cts tracking research",
        1 => "Theresa Kreckman; Jamie Engle",
    },
    10 => {
        0 => "docs ddm",
        1 => "Marc Barney; Everett Music",
    },
    11 => {
        0 => "docs read only",
        1 => "Marc Barney; Everett Music; Theresa Kreckman; Jamie Engle",
    },
    12 => {
        0 => "distwhl3rdparty",
        1 => "Theresa Kreckman; Jamie Engle",
    },
    13 => {
        0 => "ent logis b2b",
        1 => "Mark Reed; Mark Teschner",
    },
    14 => {
        0 => "ent qamasterx",
        1 => "Margaret Davis; Ron Medinger",
    },
    15 => {
        0 => "ent shipments",
        1 => "Jamie Engle; Teresa Kreckman",
    },
    16 => {
        0 => "ful distrib plan",
        1 => "Theresa Kreckman",
    },
    17 => {
        0 => "ful traffic share",
        1 => "Mark Reed; Mark Teschner",
    },
    18 => {
        0 => "ful",
        1 => "Mark Teschner",
    },
    19 => {
        0 => "hwc_international",
        1 => "Kelly Simmons",
    },
    20 => {
        0 => "masterpack/lotships",
        1 => "MBX or Distribution group",
    },
    21 => {
        0 => "medford distribution planning - mbx access",
        1 => "MBX or Distribution group",
    },
    22 => {
        0 => "nph fruit team minutes",
        1 => "MBX or Distribution group",
    },
    23 => {
        0 => "odd costco",
        1 => "Theresa Kreckman; Jamie Engle",
    },
    24 => {
        0 => "odd qvc",
        1 => "Theresa Kreckman; Jamie Engle",
    },
    25 => {
        0 => "opr ctsdata",
        1 => "Theresa Kreckman; Jamie Engle",
    },
    26 => {
        0 => "opr selectinterface",
        1 => "Jamie Engle; Theresa Kreckman",
    },
    27 => {
        0 => "opr worldship",
        1 => "Chris Larson; Jamie Engle",
    },
    28 => {
        0 => "opr-dropship",
        1 => "Theresa Kreckman; Carolyn Townsend",
    },
    29 => {
        0 => "opr-ship docks info",
        1 => "Chris Larson; Pam Saxbury",
    },
    30 => {
        0 => "opr-shipoutbol",
        1 => "Mark Reed; Mark Teschner; Theresa Kreckman",
    },
    31 => {
        0 => "proxy internet",
        1 => "Chris Works",
    },
    32 => {
        0 => "sap users",
        1 => "MBX or Distribution group",
    },
    33 => {
        0 => "trailertracking",
        1 => "Theresa Kreckman; Everett Music; Jamie Engle",
    },
    34 => {
        0 => "vendorunitaccess",
        1 => "Carolyn Townsend; Everett Music; Theresa Kreckman",
    },
    35 => {
        0 => "wms dist",
        1 => "MBX or Distribution group",
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
