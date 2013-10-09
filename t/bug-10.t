#!/usr/bin/env perl
use strict;
use warnings;
use Test::More;

use Spreadsheet::ParseXLSX;

my $wb = Spreadsheet::ParseXLSX->new->parse('t/data/bug-10.xlsx');
is($wb->worksheet_count, 4);

{
    my @contents = (
        [ ['Foo01', 0], ['Bar01', 0] ],
        [ ['Foo02', 0], ['Bar02', 0] ],
        [ ['Foo03', 0], ['Bar03', 0] ],
        [ ['Foo04', 0], ['Bar04', 0] ],
    );
    my $ws = $wb->worksheet(0);
    for my $row (0..$#contents) {
        for my $col (0..$#{ $contents[$row] }) {
            my $cell = $ws->get_cell($row, $col);
            is($cell->value, $contents[$row][$col][0]);
            is(!!$cell->is_merged, !!$contents[$row][$col][1]);
        }
    }
    is($ws->get_merged_areas, undef);
}

{
    my @contents = (
        [ ['Foo01', 0], ['Bar01', 0] ],
        [ ['Foo02', 0], ['Bar02', 0] ],
        [ ['Foo03', 0], ['Bar03', 0] ],
        [ ['Foo04', 1], ['',      1] ],
    );
    my $ws = $wb->worksheet(1);
    for my $row (0..$#contents) {
        for my $col (0..$#{ $contents[$row] }) {
            my $cell = $ws->get_cell($row, $col);
            is($cell->value, $contents[$row][$col][0]);
            is(!!$cell->is_merged, !!$contents[$row][$col][1]);
        }
    }
    is_deeply(
        $ws->get_merged_areas,
        [ [ 3, 0, 3, 1 ] ],
    );
}

{
    my @contents = (
        [ ['Foo01', 0], ['Bar01', 0] ],
        [ ['Foo02', 0], ['Bar02', 0] ],
        [ ['Foo03', 0], ['Bar03', 1] ],
        [ ['Foo04', 0], ['',      1] ],
    );
    my $ws = $wb->worksheet(2);
    for my $row (0..$#contents) {
        for my $col (0..$#{ $contents[$row] }) {
            my $cell = $ws->get_cell($row, $col);
            is($cell->value, $contents[$row][$col][0]);
            is(!!$cell->is_merged, !!$contents[$row][$col][1]);
        }
    }
    is_deeply(
        $ws->get_merged_areas,
        [ [ 2, 1, 3, 1 ] ],
    );
}

{
    my @contents = (
        [ ['Foo01', 0], ['Bar01', 0] ],
        [ ['Foo02', 0], ['Bar02', 0] ],
        [ ['Foo03', 1], ['',      1] ],
        [ ['',      1], ['',      1] ],
        [ ['Foo04', 0], ['Bar04', 1] ],
        [ ['Foo05', 0], ['',      1] ],
        [ ['Foo06', 1], ['',      1] ],
    );
    my $ws = $wb->worksheet(3);
    for my $row (0..$#contents) {
        for my $col (0..$#{ $contents[$row] }) {
            my $cell = $ws->get_cell($row, $col);
            is($cell->value, $contents[$row][$col][0]);
            is(!!$cell->is_merged, !!$contents[$row][$col][1]);
        }
    }
    is_deeply(
        $ws->get_merged_areas,
        [
            [ 2, 0, 3, 1 ],
            [ 4, 1, 5, 1 ],
            [ 6, 0, 6, 1 ],
        ],
    );
}

done_testing;
