#!/usr/bin/env perl

use Spreadsheet::ParseExcel::Utility qw(sheetRef);
use Spreadsheet::ParseXLSX;
use Test::More tests => 58;

use strict;
use warnings;

my $spreadsheet = Spreadsheet::ParseXLSX->new();
my $workbook = $spreadsheet->parse('t/data/TestHyperlinks.xlsx');
my $worksheet = $workbook->worksheet('Sheet1');

my $expected_urls = {
    'A03' => {}, # Test for cell with no hyperlink
    'A06' => {
        desc => 'http://www.example.com',
        link => 'http://www.example.com/',
    },
    'A07' => {
        desc => 'www.example.com',
        link => 'http://www.example.com/',
    },
    'A09' => {
        desc => 'file:///..\\..\\zipple.dat',
        link => '../../zipple.dat',
    },
    'A10' => {
        desc => 'ftp://user:pass@example.net/pub/manuals/Excel.doc',
        link => 'ftp://user:pass@example.net/pub/manuals/Excel.doc',
    },

    'B6' => {
        desc => 'http://www.example.com#foo',
        link => 'http://www.example.com/#foo',
    },
    'B7' => {
        desc => 'www.example.com#foo',
        link => 'http://www.example.com/#foo',
    },

    'C6' => {
        desc => 'file:///c:\\nodir\\nofile.txt',
        link => 'file:///c:\\nodir\\nofile.txt',
    },
    'C7' => {
        desc => 'c:\\nodir\\nofile.txt',
        link => 'file:///c:\\nodir\\nofile.txt',
    },

    'D6' => {
        desc => '\\\\server\\quirks\\sometest.bat',
        link => 'file:///\\\\server\\quirks\\sometest.bat',
    },
    'D7' => {
        desc => 'SMB Link Sometest.bat',
        link => 'file:///\\\\server\\quirks\\sometest.bat',
    },

    'F7' => {
        desc => 'mailto:fred@example.net',
        link => 'mailto:fred@example.net',
    },
};

foreach my $test_cell (sort keys %$expected_urls) {
    # First check our cell reference is valid
    my ($row, $column) = sheetRef($test_cell);

    unless (defined($row) && defined($column)) {
        warn(Data::Dumper::Dumper($test_cell, $row, $column, sheetRef($test_cell)));
        die('Unable to parse cell reference: ' . $test_cell);
    }

    # Now extract out our expected data
    my $link = $expected_urls->{$test_cell}->{link};
    my $desc = $expected_urls->{$test_cell}->{desc};
    my $relative_link = $expected_urls->{$test_cell}->{rel} || 0;

    if ($relative_link) {
        $link = 'file:///t/data/' . $link;
    }

    # Tidy up our cell reference as I frigged some of them to ensure test order
    if ($test_cell =~ /^([A-Z]{1})0(\d{1})$/) {
        $test_cell = $1 . $2;
    }

    my $cell = $worksheet->get_cell($row, $column);
    ok(defined($cell), sprintf('Cell "%s" defined', $test_cell));

    SKIP: {
        skip(sprintf('Cell "%s" not defined', $test_cell), ($link ? 3 : 1)) unless(defined($cell));

        my $hyperlink = $cell->get_hyperlink();

        if ($link) {
            ok(ref($hyperlink) eq 'ARRAY', 'Got hyperlink information from cell: ' . $test_cell);
            is($hyperlink->[0], $desc, sprintf('Cell "%s" hyperlink description matches "%s"', $test_cell, $desc));
            is($hyperlink->[1], $link, sprintf('Cell "%s" hyperlink destination matches "%s"', $test_cell, $link));
        } else {
            is($hyperlink, undef, sprintf('Cell "%s" has no hyperlink', $test_cell));
        }
    }
}

# The following tests return different values to Spreadsheet::ParseExcel - I don't know if this is correct or not
my $todo_expected_urls = {
    'E6' => {
        desc => 'TestHyperlinks.xlsx',
        rel  => 1,
        link => 'TestHyperlinks.xlsx',
    },
    'E7' => {
        desc => 'Rel: TestHyperlinks.xlsx',
        rel  => 1,
        link => 'TestHyperlinks.xlsx',
    },

    'F6' => {
        desc => 'Sheet1!A7',
        link => '#Sheet1%21A7',
    },
};

TODO: {
    local $TODO = 'Confirmation of difference to Spreadsheet::ParseExcel required';

    foreach my $test_cell (sort keys %$todo_expected_urls) {
        # First check our cell reference is valid
        my ($row, $column) = sheetRef($test_cell);

        unless (defined($row) && defined($column)) {
            warn(Data::Dumper::Dumper($test_cell, $row, $column, sheetRef($test_cell)));
            die('Unable to parse cell reference: ' . $test_cell);
        }

        # Now extract out our expected data
        my $link = $todo_expected_urls->{$test_cell}->{link};
        my $desc = $todo_expected_urls->{$test_cell}->{desc};
        my $relative_link = $todo_expected_urls->{$test_cell}->{rel} || 0;

        if ($relative_link) {
            $link = 'file:///t/data/' . $link;
        }

        # Tidy up our cell reference as I frigged some of them to ensure test order
        if ($test_cell =~ /^([A-Z]{1})0(\d{1})$/) {
            $test_cell = $1 . $2;
        }

        my $cell = $worksheet->get_cell($row, $column);
        ok(defined($cell), sprintf('Cell "%s" defined', $test_cell));

        SKIP: {
            skip(sprintf('Cell "%s" not defined', $test_cell), ($link ? 3 : 1)) unless(defined($cell));

            my $hyperlink = $cell->get_hyperlink();

            if ($link) {
                ok(ref($hyperlink) eq 'ARRAY', 'Got hyperlink information from cell: ' . $test_cell);
                is($hyperlink->[0], $desc, sprintf('Cell "%s" hyperlink description matches "%s"', $test_cell, $desc));
                is($hyperlink->[1], $link, sprintf('Cell "%s" hyperlink destination matches "%s"', $test_cell, $link));
            } else {
                is($hyperlink, undef, sprintf('Cell "%s" has no hyperlink', $test_cell));
            }
        }
    }
}

exit;
