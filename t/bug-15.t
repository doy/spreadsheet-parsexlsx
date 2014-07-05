#!/usr/bin/env perl
use strict;
use warnings;
use Test::More;

use Spreadsheet::ParseXLSX;

my $wb = Spreadsheet::ParseXLSX->new->parse('t/data/bug-15.xlsx');
my $ws = $wb->worksheet(1);

my $b2 = $ws->get_cell(1, 1);
ok(exists $b2->get_format->{Hidden});
ok(exists $b2->get_format->{Lock});
ok($b2->get_format->{IgnoreProtection});
ok(!$b2->get_format->{Hidden});
ok($b2->get_format->{Lock});

my $b3 = $ws->get_cell(2, 1);
ok(exists $b3->get_format->{Hidden});
ok(exists $b3->get_format->{Lock});
ok(!$b3->get_format->{IgnoreProtection});
ok(!$b3->get_format->{Hidden});
ok(!$b3->get_format->{Lock});

my $b4 = $ws->get_cell(3, 1);
ok(exists $b4->get_format->{Hidden});
ok(exists $b4->get_format->{Lock});
ok(!$b4->get_format->{IgnoreProtection});
ok($b4->get_format->{Hidden});
ok(!$b4->get_format->{Lock});

done_testing;
