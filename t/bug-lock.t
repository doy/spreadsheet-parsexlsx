#!/usr/bin/env perl
use strict;
use warnings;
use Test::More;

use Spreadsheet::ParseXLSX;

my $wb = Spreadsheet::ParseXLSX->new->parse('t/data/bug-lock.xlsx');
my $ws = $wb->worksheet(0);

my $b1 = $ws->get_cell(0, 0);
ok($b1->get_format->{Lock});

my $b2 = $ws->get_cell(1, 0);
ok(!$b2->get_format->{Lock});

my $b3 = $ws->get_cell(2, 0);
ok($b3->get_format->{Lock});

my $b4 = $ws->get_cell(3, 0);
ok(!$b4->get_format->{Lock});

my $b5 = $ws->get_cell(4, 0);
ok($b5->get_format->{Lock});

done_testing;
