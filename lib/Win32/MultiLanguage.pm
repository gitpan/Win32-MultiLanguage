# MultiLanguage.pm -- ...
#
# $Id$

package Win32::MultiLanguage;
use 5.008;
use strict;
use warnings;

# MLDETECTCP
use constant MLDETECTCP_NONE              => 0x00000000;
use constant MLDETECTCP_7BIT              => 0x00000001;
use constant MLDETECTCP_8BIT              => 0x00000002;
use constant MLDETECTCP_DBCS              => 0x00000004;
use constant MLDETECTCP_HTML              => 0x00000008;

# MIMECONTF
use constant MIMECONTF_MAILNEWS           => 0x00000001;
use constant MIMECONTF_BROWSER            => 0x00000002;
use constant MIMECONTF_MINIMAL            => 0x00000004;
use constant MIMECONTF_IMPORT             => 0x00000008;
use constant MIMECONTF_SAVABLE_MAILNEWS   => 0x00000100;
use constant MIMECONTF_SAVABLE_BROWSER    => 0x00000200;
use constant MIMECONTF_EXPORT             => 0x00000400;
use constant MIMECONTF_PRIVCONVERTER      => 0x00010000;
use constant MIMECONTF_VALID              => 0x00020000;
use constant MIMECONTF_VALID_NLS          => 0x00040000;
use constant MIMECONTF_MIME_IE4           => 0x10000000;
use constant MIMECONTF_MIME_LATEST        => 0x20000000;
use constant MIMECONTF_MIME_REGISTRY      => 0x40000000;

# MLDETECTF
use constant MLDETECTF_MAILNEWS           => 0x00000001;
use constant MLDETECTF_BROWSER            => 0x00000002;
use constant MLDETECTF_VALID              => 0x00000004;
use constant MLDETECTF_VALID_NLS          => 0x00000008;
use constant MLDETECTF_PRESERVE_ORDER     => 0x00000010;
use constant MLDETECTF_PREFERRED_ONLY     => 0x00000020;
use constant MLDETECTF_FILTER_SPECIALCHAR => 0x00000040;

our $VERSION = '0.05';

require XSLoader;
XSLoader::load('Win32::MultiLanguage', $VERSION);

# Preloaded methods go here.

1;

__END__

=pod

=head1 NAME

Win32::MultiLanguage - Interface to IMultiLanguage I18N routines

=head1 SYNOPSIS

  use Win32::MultiLanguage;
  # @@

=head1 DESCRIPTION

Win32::MultiLanguage is an experimental wrapper module for the
Windows IMultiLanguage interfaces that comes with Internet
Explorer version 4 and later. Mlang.dll implements routines for
dealing with character encodings, code pages, and locales.

=head1 CONSTANTS

=head2 MLDETECTCP

=over 2

=item MLDETECTCP_NONE

Default setting will be used.

=item MLDETECTCP_7BIT

Input stream consists of 7-bit data.

=item MLDETECTCP_8BIT

Input stream consists of 8-bit data.

=item MLDETECTCP_DBCS

Input stream consists of double-byte data.

=item MLDETECTCP_HTML

Input stream is an HTML page.

=back

=head2 MIMECONTF

=over 2

=item MIMECONTF_MAILNEWS

Code page is meant to display on mail and news clients. 

=item MIMECONTF_BROWSER

Code page is meant to display on browser clients. 

=item MIMECONTF_MINIMAL

Code page is meant to display in minimal view. This value is generally
not used. 

=item MIMECONTF_IMPORT

Value that indicates that all of the import code pages should be enumerated. 

=item MIMECONTF_SAVABLE_MAILNEWS

Code page includes encodings for mail and news clients to save a document in. 

=item MIMECONTF_SAVABLE_BROWSER

Code page includes encodings for browser clients to save a document
in. 

=item MIMECONTF_EXPORT

Value that indicates that all of the export code pages should be
enumerated. 

=item MIMECONTF_PRIVCONVERTER

Value that indicates the encoding requires (or has) a private
conversion engine. A client of IEnumCodePage doesn't use this value. 

=item MIMECONTF_VALID

Value that indicates the corresponding encoding is supported on the
system. 

=item MIMECONTF_VALID_NLS

Value that indicates that only the language support file should be
validated. Normally, both the language support file and the
supporting font are checked. 

=item MIMECONTF_MIME_IE4

Value that indicates the Microsoft® Internet Explorer 4.0 MIME data
from MLang's internal data should be used. 

=item MIMECONTF_MIME_LATEST

Value that indicates that the latest MIME data from MLang's internal
data should be used. 

=item MIMECONTF_MIME_REGISTRY

Value that indicates that the MIME data stored in the registry should
be used. 

=back

=head2 MLDETECTF

=over 2

=item MLDETECTF_MAILNEWS

Not currently supported.

=item MLDETECTF_BROWSER

Not currently supported.

=item MLDETECTF_VALID

Detection result must be valid for conversion and text rendering.

=item MLDETECTF_VALID_NLS

Detection result must be valid for conversion.

=item MLDETECTF_PRESERVE_ORDER

Preserve preferred code page order. This is meaningful only if you have
set the @@puiPreferredCodePages parameter in C<DetectOutboundCodePage>.

=item MLDETECTF_PREFERRED_ONLY

Only return one of the preferred code pages as the detection result.
This is meaningful only if you have set the @@puiPreferredCodePages
parameter in C<DetectOutboundCodePage>.

=item MLDETECTF_FILTER_SPECIALCHAR

Filter out graphical symbols and punctuation.

=back

=head1 ROUTINES

=over 4

=item DetectInputCodepage($octets [, $flags [, $codepage]])

Detects the code page of the given string $octets. An optional
$flags parameter may be specified, a combination of C<MLDETECTCP>
constants as defined above, if not specified C<MLDETECTCP_NONE>
will be used as default. An optional $codepage can also be
specified, if this value is set to zero, this API returns all
possible encodings. Otherwise, it lists only those encodings
related to this parameter. The default is zero.

It will return a reference to an array of hash references of
which each represents a C<DetectEncodingInfo> strucure with the
following keys

  LangID     => ..., # primary language identifier
  CodePage   => ..., # detected Win32-defined code page
  DocPercent => ..., # Percentage in the detected language
  Confidence => ..., # degree to which the detected data is correct

See L<http://msdn.microsoft.com/workshop/misc/mlang/reference/structures/detectencodinginfo.asp>
for details.

=item GetCodePageInfo($codepage, $langid)

...

=item GetCodePageDescription($codepage, $locale)

...

=item GetRfc1766FromLcid($locale)

...

=item DetectOutboundCodePage($utf8 [, $flags [, \@cp ]])

...

=item GetCharsetInfo($charset)

...

=back

=head1 IMPLEMENTATION STATUS

  Legend:

    + means is implemented
    ? means might get implemented
    - means unlikely that this gets implemented

  IMultiLanguage
    + GetCharsetInfo
    + GetRfc1766FromLcid
    ? GetNumberOfCodePageInfo
    ? GetFamilyCodePage
    ? EnumCodePages
    ? IsConvertible
    ? ConvertString
    ? GetLcidFromRfc1766
    ? EnumRfc1766
    ? GetRfc1766Info
    - ConvertStringToUnicode
    - ConvertStringFromUnicode
    - ConvertStringReset
    - CreateConvertCharset
    
  IMultiLanguage2
    + GetCodePageInfo
    + DetectInputCodepage
    + GetCodePageDescription
    ? ValidateCodePage
    ? IsCodePageInstallable
    ? GetNumberOfScripts
    ? EnumScripts
    ? ValidateCodePageEx
    - ConvertStringInIStream
    - ConvertStringToUnicodeEx
    - ConvertStringFromUnicodeEx
    - DetectCodepageInIStream
    - SetMimeDBSource
    
  IMultiLanguage3
    + DetectOutboundCodePage
    - DetectOutboundCodePageInIStream

=head1 SEE ALSO

=over 4

=item * L<http://msdn.microsoft.com/workshop/misc/mlang/mlang.asp>

=back

=head1 WARNING

This is pre-alpha software.

=head1 AUTHOR AND COPYRIGHT

Copyright (C) 2004 by Bjoern Hoehrmann E<lt>bjoern@hoehrmann.deE<gt>.

This library is free software; you can redistribute it and/or modify
it under the same terms as Perl itself, either Perl version 5.8.2 or,
at your option, any later version of Perl 5 you may have available.

=cut
