# NAME

Spreadsheet::ParseXLSX - parse XLSX files

# VERSION

version 0.28

# SYNOPSIS

```perl
use Spreadsheet::ParseXLSX;

my $parser = Spreadsheet::ParseXLSX->new;
my $workbook = $parser->parse("file.xlsx");
# see Spreadsheet::ParseExcel for further documentation
```

# DESCRIPTION

This module is an adaptor for [Spreadsheet::ParseExcel](https://metacpan.org/pod/Spreadsheet%3A%3AParseExcel) that reads XLSX files.
For documentation about the various data that you can retrieve from these
classes, please see [Spreadsheet::ParseExcel](https://metacpan.org/pod/Spreadsheet%3A%3AParseExcel),
[Spreadsheet::ParseExcel::Workbook](https://metacpan.org/pod/Spreadsheet%3A%3AParseExcel%3A%3AWorkbook), [Spreadsheet::ParseExcel::Worksheet](https://metacpan.org/pod/Spreadsheet%3A%3AParseExcel%3A%3AWorksheet),
and [Spreadsheet::ParseExcel::Cell](https://metacpan.org/pod/Spreadsheet%3A%3AParseExcel%3A%3ACell).

# METHODS

## new(%opts)

Returns a new parser instance. Takes a hash of parameters:

- Password

    Password to use for decrypting encrypted files.

## parse($file, $formatter)

Parses an XLSX file. Parsing errors throw an exception. `$file` can be either
a filename or an open filehandle. Returns a
[Spreadsheet::ParseExcel::Workbook](https://metacpan.org/pod/Spreadsheet%3A%3AParseExcel%3A%3AWorkbook) instance containing the parsed data.
The `$formatter` argument is an optional formatter class as described in [Spreadsheet::ParseExcel](https://metacpan.org/pod/Spreadsheet%3A%3AParseExcel).

# INCOMPATIBILITIES

This module returns data using classes from [Spreadsheet::ParseExcel](https://metacpan.org/pod/Spreadsheet%3A%3AParseExcel), so for
the most part, it should just be a drop-in replacement. That said, there are a
couple areas where the data returned is intentionally different:

- Colors

    In Spreadsheet::ParseExcel, colors are represented by integers which index into
    the color table, and you have to use
    `Spreadsheet::ParseExcel->ColorIdxToRGB` in order to get the actual value
    out. In Spreadsheet::ParseXLSX, while the color table still exists, cells are
    also allowed to specify their color directly rather than going through the
    color table. In order to avoid confusion, I normalize all color values in
    Spreadsheet::ParseXLSX to their string RGB format (`"#0088ff"`). This affects
    the `Fill`, `BdrColor`, and `BdrDiag` properties of formats, and the
    `Color` property of fonts. Note that the default color is represented by
    `undef` (the same thing that `ColorIdxToRGB` would return).

- Formulas

    Spreadsheet::ParseExcel doesn't support formulas. Spreadsheet::ParseXLSX
    provides basic formula support by returning the text of the formula as part of
    the cell data. You can access it via `$cell->{Formula}`. Note that the
    restriction still holds that formula cell values aren't available unless they
    were explicitly provided when the spreadsheet was written.

# BUGS

- Large spreadsheets may cause segfaults on perl 5.14 and earlier

    This module internally uses XML::Twig, which makes it potentially subject to
    [Bug #71636 for XML-Twig: Segfault with medium-sized document](https://rt.cpan.org/Public/Bug/Display.html?id=71636)
    on perl versions 5.14 and below (the underlying bug with perl weak references
    was fixed in perl 5.15.5). The larger and more complex the spreadsheet, the
    more likely to be affected, but the actual size at which it segfaults is
    platform dependent. On a 64-bit perl with 7.6gb memory, it was seen on
    spreadsheets about 300mb and above. You can work around this adding
    `XML::Twig::_set_weakrefs(0)` to your code before parsing the spreadsheet,
    although this may have other consequences such as memory leaks.

- Worksheets without the `dimension` tag are not supported
- Intra-cell formatting is discarded
- Shared formulas are not supported

    Shared formula support will require an actual formula parser and quite a bit of
    custom logic, since the only thing stored in the document is the formula for
    the base cell - updating the cell references in the formulas in the rest of the
    cells is handled by the application. Values for these cells are still handled
    properly.

In addition, there are still a few areas which are not yet implemented (the
XLSX spec is quite large). If you run into any of those, bug reports are quite
welcome.

Please report any bugs to GitHub Issues at
[https://github.com/MichaelDaum/spreadsheet-parsexlsx/issues](https://github.com/MichaelDaum/spreadsheet-parsexlsx/issues).

# SEE ALSO

[Spreadsheet::ParseExcel](https://metacpan.org/pod/Spreadsheet%3A%3AParseExcel): The equivalent, for XLS files.

[Spreadsheet::XLSX](https://metacpan.org/pod/Spreadsheet%3A%3AXLSX): An older, less robust and featureful implementation.

# SUPPORT

You can find this documentation for this module with the perldoc command.

```
perldoc Spreadsheet::ParseXLSX
```

You can also look for information at:

- MetaCPAN

    [https://metacpan.org/release/Spreadsheet-ParseXLSX](https://metacpan.org/release/Spreadsheet-ParseXLSX)

- RT: CPAN's request tracker

    [http://rt.cpan.org/NoAuth/Bugs.html?Dist=Spreadsheet-ParseXLSX](http://rt.cpan.org/NoAuth/Bugs.html?Dist=Spreadsheet-ParseXLSX)

- Github

    [https://github.com/MichaelDaum/spreadsheet-parsexlsx](https://github.com/MichaelDaum/spreadsheet-parsexlsx)

- CPAN Ratings

    [http://cpanratings.perl.org/d/Spreadsheet-ParseXLSX](http://cpanratings.perl.org/d/Spreadsheet-ParseXLSX)

# SPONSORS

Parts of this code were paid for by

- Socialflow [http://socialflow.com](http://socialflow.com)

# AUTHOR

Jesse Luehrs <doy@tozt.net>

# CONTRIBUTORS

- Alexey Mazurin <mazurin.alexey@gmail.com>
- Dave Clarke &lt;david\_clarke@verizon.net>
- Fitz Elliott <felliott@fiskur.org>
- FL <f20@reckon.co.uk>
- Meredith Howard <mhoward@roomag.org>
- MichaelDaum <daum@michaeldaumconsulting.com>
- rdboisvert <rdbprog@gmail.com>
- Ricardo Signes <rjbs@cpan.org>
- Robert Boisvert &lt;robert.boisvert@PABET-J069H12.sncrcorp.net>
- Steve Simms <steve@deefs.net>
- Stuart Watt <stuart@morungos.com>
- zhouzhen1 <zhouzhen1@gmail.com>

# COPYRIGHT AND LICENSE

This software is Copyright (c) 2024 by Jesse Luehrs.

This is free software, licensed under:

```
The MIT (X11) License
```
