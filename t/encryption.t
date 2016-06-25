#!/usr/bin/env perl
use strict;
use warnings;
use Test::More;

use Spreadsheet::ParseXLSX;

{
    my $filename = 't/data/encryption-agile-123q.xlsx';
    my @inputs = (
        $filename,
        do { open my $fh, '<', $filename or die; $fh },
        do { open my $fh, '<', $filename or die; local $/; my $d = <$fh>; \$d },
    );

    my $parser = Spreadsheet::ParseXLSX->new(Password => '123q');
    for my $input (@inputs) {
        my $workbook = $parser->parse($input);

        my $worksheet = $workbook->worksheet(0);
        my $cell = $worksheet->get_cell(1, 1);
        is($cell->value, 'abcdefgABCDEFG');
    }
}

{
    my $filename = 't/data/encryption-standard-default-password.xlsx';
    my @inputs = (
        $filename,
        do { open my $fh, '<', $filename or die; $fh },
        do { open my $fh, '<', $filename or die; local $/; my $d = <$fh>; \$d },
    );

    my $parser = Spreadsheet::ParseXLSX->new(Password => '');
    for my $input (@inputs) {
        my $workbook = $parser->parse($input);

        my $worksheet = $workbook->worksheet(0);
        my $cell = $worksheet->get_cell(22, 8);
        is($cell->value, 1911);
    }
}

done_testing;
