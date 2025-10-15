<?php
namespace Composer\Pcre;

/**
 * Minimal polyfill for composer/pcre Preg wrapper used by PhpSpreadsheet.
 * This provides only the static methods needed by the library.
 */
final class Preg
{
    public static function isMatch($pattern, $subject, $flags = 0)
    {
        return (bool)preg_match($pattern, $subject);
    }

    public static function match($pattern, $subject, &$matches = null, $flags = 0, $offset = 0)
    {
        return preg_match($pattern, $subject, $matches, $flags, $offset);
    }

    public static function matchAll($pattern, $subject, &$matches = null, $flags = 0, $offset = 0)
    {
        return preg_match_all($pattern, $subject, $matches, $flags, $offset);
    }

    public static function replace($pattern, $replacement, $subject, $limit = -1, &$count = null)
    {
        return preg_replace($pattern, $replacement, $subject, $limit, $count);
    }

    public static function replaceCallback($pattern, callable $callback, $subject, $limit = -1, &$count = null)
    {
        return preg_replace_callback($pattern, $callback, $subject, $limit, $count);
    }

    public static function split($pattern, $subject, $limit = -1, $flags = 0)
    {
        return preg_split($pattern, $subject, $limit, $flags);
    }
}

