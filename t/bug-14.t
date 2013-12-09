#!/usr/bin/env perl
use strict;
use warnings;
use Test::More;

use Spreadsheet::ParseXLSX;

my $wb = Spreadsheet::ParseXLSX->new->parse('t/data/bug-14.xlsx');
my $ws = $wb->worksheet(0);

for my $row (0..6) {
    for my $col ($row..6) {
        next if $row == 5 && $col == 6;

        my $font = $ws->get_cell($row, $col)->get_format->{Font};
        is($font->{Name}, 'Arial');
        is(!!$font->{Bold}, $row == 1 || $col == 1);
        is(!!$font->{Italic}, $row == 2 || $col == 2);
        is($font->{Height}, 10);
        is(!!$font->{Underline}, $row == 3 || $col == 3);
        if ($row == 3 || $col == 3) {
            is($font->{UnderlineStyle}, 1);
        }
        is($font->{Color}, '#000000');
        is(!!$font->{Strikeout}, $row == 4 || $col == 4);
        is(
            $font->{Super},
            $row == 5 || $col == 5 ? 2
          : $row == 6 || $col == 6 ? 1
          :                          0
        );
    }
}

done_testing;
