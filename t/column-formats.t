#!/usr/bin/env perl
use strict;
use warnings;
use Test::More;

use Spreadsheet::ParseXLSX;

my $wb = Spreadsheet::ParseXLSX->new->parse('t/data/column-formats.xlsx');
my $ws = $wb->worksheet(0);

ok(my $col_format_nos = $ws->{ColFmtNo});

my @col_formats = map { $wb->{Format}[ $_ ] } @$col_format_nos;
is_deeply($col_formats[0]->{Fill}, [1, '#FF0000', '#FFFFFF']);

is($col_formats[1]->{AlignH}, 3);
is($col_formats[1]->{AlignV}, 0);

my $font = $col_formats[2]->{Font};
is_deeply($font->{Name}, 'Arial');
is_deeply($font->{Height}, 16);
is_deeply($font->{Bold}, 1);


done_testing;
