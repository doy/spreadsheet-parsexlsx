#!/usr/bin/env perl
use strict;
use warnings;
use Test::More;

use Spreadsheet::ParseXLSX;

my $wb = Spreadsheet::ParseXLSX->new->parse('t/data/bug-11.xlsx');
is($wb->worksheet_count, 1);

my $ws = $wb->worksheet(0);
is($ws->get_cell(0, 0)->value, "foobarbaz");
my $rich_text_data = $ws->get_cell(0, 0)->get_rich_text;
is($rich_text_data->[0][0], 0);
ok(!$rich_text_data->[0][1]->{Italic});
is($rich_text_data->[1][0], 3);
ok($rich_text_data->[1][1]->{Italic});
is($rich_text_data->[2][0], 6);
ok(!$rich_text_data->[2][1]->{Italic});

done_testing;
