#!/usr/bin/env perl
use strict;
use warnings;
use Test::More;
use utf8;

use Spreadsheet::ParseXLSX;
use Data::Dumper;

my $wb = Spreadsheet::ParseXLSX->new->parse('t/data/bug-72.xlsx');
my $ws = $wb->worksheet(0);

my $b1 = $ws->get_cell(0, 0);
my $b2 = $ws->get_cell(1, 0);

is $b1->value(), "日本語あいうえお";
is $b2->value(), "日本語あいうえお";

done_testing;
