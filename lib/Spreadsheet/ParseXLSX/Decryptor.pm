package Spreadsheet::ParseXLSX::Decryptor;
use strict;
use warnings;

use Crypt::Mode::CBC;
use Crypt::Mode::ECB;
use Digest::SHA ();
use Encode ();
use File::Temp 'tempfile';
use MIME::Base64 ();
use OLE::Storage_Lite;

use Spreadsheet::ParseXLSX::Decryptor::Standard;
use Spreadsheet::ParseXLSX::Decryptor::Agile;

sub open {
    my $class = shift;

    my ($filename, $password) = @_;

    $password = $password || 'VelvetSweatshop';

    my ($infoFile, $packageFile) = _getCompoundData($filename, ['EncryptionInfo', 'EncryptedPackage']);

    my $xlsx;

    eval {
        my $infoFH = IO::File->new();
        $infoFH->open($infoFile);
        $infoFH->binmode();

        my $buffer;
        $infoFH->read($buffer, 8);
        my ($majorVers, $minorVers) = unpack('SS', $buffer);

        if ($majorVers == 4 && $minorVers == 4) {
            $xlsx = agileDecryption($infoFH, $packageFile, $password);
        } else {
            $xlsx = standardDecryption($infoFH, $packageFile, $password);
        }
        $infoFH->close();
    };
    unlink $infoFile, $packageFile;
    die $@ if $@;

    return $xlsx;
}

sub _getCompoundData {
    my $filename = shift;
    my $names = shift;

    my @files;

    my $storage = OLE::Storage_Lite->new($filename);

    foreach my $name (@{$names}) {
        my @data = $storage->getPpsSearch([OLE::Storage_Lite::Asc2Ucs($name)], 1, 1);
        if ($#data < 0) {
            push @files, undef;
        } else {
            my ($fh, $filename) = File::Temp::tempfile();
            my $out = IO::Handle->new_from_fd($fh, 'w') || die "TempFile error!";
            $out->write($data[0]->{Data});
            $out->close();
            push @files, $filename;
        }
    }

    return @files;
}

sub standardDecryption {
    my ($infoFH, $packageFile, $password) = @_;

    my $buffer;
    my $n = $infoFH->read($buffer, 24);

    my ($encryptionHeaderSize, undef, undef, $algID, $algIDHash, $keyBits) = unpack('LLLLLL', $buffer);

    $infoFH->seek($encryptionHeaderSize - 0x14, IO::File::SEEK_CUR);

    $infoFH->read($buffer, 4);

    my $saltSize = unpack('L', $buffer);

    my ($salt, $encryptedVerifier, $verifierHashSize, $encryptedVerifierHash);

    $infoFH->read($salt, 16);
    $infoFH->read($encryptedVerifier, 16);

    $infoFH->read($buffer, 4);
    $verifierHashSize = unpack('L', $buffer);

    $infoFH->read($encryptedVerifierHash, 32);
    $infoFH->close();

    my ($cipherAlgorithm, $hashAlgorithm);

    if ($algID == 0x0000660E || $algID == 0x0000660F || $algID == 0x0000660E) {
        $cipherAlgorithm = 'AES';
    } else {
        die sprintf('Unsupported encryption algorithm: 0x%.8x', $algID);
    }

    if ($algIDHash == 0x00008004) {
        $hashAlgorithm = 'SHA-1';
    } else {
        die sprintf('Unsupported hash algorithm: 0x%.8x', $algIDHash);
    }

    my $decryptor = Spreadsheet::ParseXLSX::Decryptor::Standard->new({
                  cipherAlgorithm => $cipherAlgorithm,
                  cipherChaining  => 'ECB',
                  hashAlgorithm   => $hashAlgorithm,
                  salt            => $salt,
                  password        => $password,
                  keyBits         => $keyBits,
                  spinCount       => 50000
              });

    $decryptor->verifyPassword($encryptedVerifier, $encryptedVerifierHash);

    my $in = new IO::File;
    $in->open("<$packageFile") || die 'File/handle opening error';
    $in->binmode();

    my ($fh, $filename) = File::Temp::tempfile();
    binmode($fh);
    my $out = IO::Handle->new_from_fd($fh, 'w') || die "TempFile error!";

    my $inbuf;
    $in->read($inbuf, 8);
    my $fileSize = unpack('L', $inbuf);

    $decryptor->decryptFile($in, $out, 1024, $fileSize);

    $in->close();
    $out->close();

    return $filename;
}

sub agileDecryption {
    my ($infoFH, $packageFile, $password) = @_;

    my $xml = XML::Twig->new;
    $xml->parse($infoFH);

    my ($info) = $xml->find_nodes('//encryption/keyEncryptors/keyEncryptor/p:encryptedKey');

    my $encryptedVerifierHashInput = MIME::Base64::decode($info->att('encryptedVerifierHashInput'));
    my $encryptedVerifierHashValue = MIME::Base64::decode($info->att('encryptedVerifierHashValue'));
    my $encryptedKeyValue = MIME::Base64::decode($info->att('encryptedKeyValue'));

    my $keyDecryptor = Spreadsheet::ParseXLSX::Decryptor::Agile->new({
                  cipherAlgorithm => $info->att('cipherAlgorithm'),
                  cipherChaining  => $info->att('cipherChaining'),
                  hashAlgorithm   => $info->att('hashAlgorithm'),
                  salt            => MIME::Base64::decode($info->att('saltValue')),
                  password        => $password,
                  keyBits         => 0 + $info->att('keyBits'),
                  spinCount       => 0 + $info->att('spinCount'),
                  blockSize       => 0 + $info->att('blockSize')
              });

    $keyDecryptor->verifyPassword($encryptedVerifierHashInput, $encryptedVerifierHashValue);

    my $key = $keyDecryptor->decrypt($encryptedKeyValue, "\x14\x6e\x0b\xe7\xab\xac\xd0\xd6");

    ($info) = $xml->find_nodes('//encryption/keyData');

    my $fileDecryptor = Spreadsheet::ParseXLSX::Decryptor::Agile->new({
                  cipherAlgorithm => $info->att('cipherAlgorithm'),
                  cipherChaining  => $info->att('cipherChaining'),
                  hashAlgorithm   => $info->att('hashAlgorithm'),
                  salt            => MIME::Base64::decode($info->att('saltValue')),
                  password        => $password,
                  keyBits         => 0 + $info->att('keyBits'),
                  blockSize       => 0 + $info->att('blockSize')
              });

    my $in = new IO::File;
    $in->open("<$packageFile") || die 'File/handle opening error';
    $in->binmode();

    my ($fh, $filename) = File::Temp::tempfile();
    binmode($fh);
    my $out = IO::Handle->new_from_fd($fh, 'w') || die "TempFile error!";

    my $inbuf;
    $in->read($inbuf, 8);
    my $fileSize = unpack('L', $inbuf);

    $fileDecryptor->decryptFile($in, $out, 4096, $key, $fileSize);

    $in->close();
    $out->close();

    return $filename;
}

sub new {
    my $class = shift;
    my $self = shift;

    $self->{keyLength} = $self->{keyBits} / 8;

    if ($self->{hashAlgorithm} eq 'SHA512') {
        $self->{hashProc} = \&Digest::SHA::sha512;
    } elsif ($self->{hashAlgorithm} eq 'SHA-1') {
        $self->{hashProc} = \&Digest::SHA::sha1;
    } elsif ($self->{hashAlgorithm} eq 'SHA256') {
        $self->{hashProc} = \&Digest::SHA::sha256;
    } else {
        die "Unsupported hash algorithm: $self->{hashAlgorithm}";
    }

    return bless $self, $class;
}

1;