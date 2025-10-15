<?php

namespace PhpOffice\PhpSpreadsheet;

/**
 * Spreadsheet theme information covering colours and fonts used by the
 * workbook. The Reader populates the properties when a document supplies a
 * custom theme and the Writer consumes them when recreating the theme part of
 * the XLSX file.
 */
class Theme
{
    public const HYPERLINK_THEME = 10;
    public const FOLLOWED_HYPERLINK_THEME = 11;

    private const DEFAULT_THEME_COLOR_NAME = 'Office';
    private const DEFAULT_THEME_FONT_NAME = 'Office';

    /**
     * Order of colours as defined by the OOXML specification. The numeric index
     * is used when colours reference the theme by index instead of by name.
     */
    private const THEME_COLOUR_ORDER = [
        'lt1',
        'dk1',
        'lt2',
        'dk2',
        'accent1',
        'accent2',
        'accent3',
        'accent4',
        'accent5',
        'accent6',
        'hlink',
        'folHlink',
    ];

    /** @var array<string, string> */
    private const DEFAULT_THEME_COLOURS = [
        'lt1' => 'FFFFFF',
        'dk1' => '000000',
        'lt2' => 'EEECE1',
        'dk2' => '1F497D',
        'accent1' => '4F81BD',
        'accent2' => 'C0504D',
        'accent3' => '9BBB59',
        'accent4' => '8064A2',
        'accent5' => '4BACC6',
        'accent6' => 'F79646',
        'hlink' => '0000FF',
        'folHlink' => '800080',
    ];

    private string $themeColorName = self::DEFAULT_THEME_COLOR_NAME;

    private string $themeFontName = self::DEFAULT_THEME_FONT_NAME;

    /** @var array<string, string> */
    private array $themeColours = self::DEFAULT_THEME_COLOURS;

    /** @var array<int, string> */
    private array $themeColoursByIndex = [];

    /** @var array<string, int> */
    private array $colourIndexByName = [];

    private string $majorFontLatin = 'Cambria';

    private string $majorFontEastAsian = '';

    private string $majorFontComplexScript = 'Times New Roman';

    /** @var array<string, string> */
    private array $majorFontSubstitutions = [];

    private string $minorFontLatin = 'Calibri';

    private string $minorFontEastAsian = '';

    private string $minorFontComplexScript = 'Times New Roman';

    /** @var array<string, string> */
    private array $minorFontSubstitutions = [];

    public function __construct()
    {
        $this->initialiseColourIndexes();
    }

    private function initialiseColourIndexes(): void
    {
        $this->themeColoursByIndex = [];
        $this->colourIndexByName = [];
        foreach (self::THEME_COLOUR_ORDER as $index => $colourName) {
            if (!isset($this->themeColours[$colourName])) {
                continue;
            }
            $this->themeColoursByIndex[$index] = $this->themeColours[$colourName];
            $this->colourIndexByName[$colourName] = $index;
        }
    }

    public function getThemeColorName(): string
    {
        return $this->themeColorName;
    }

    public function setThemeColorName(string $themeColorName): void
    {
        $this->themeColorName = $themeColorName;
    }

    public function getThemeFontName(): string
    {
        return $this->themeFontName;
    }

    public function setThemeFontName(string $themeFontName): void
    {
        $this->themeFontName = $themeFontName;
    }

    /**
     * @return array<string, string>
     */
    public function getThemeColors(): array
    {
        return $this->themeColours;
    }

    public function getColourByIndex(int $index): ?string
    {
        return $this->themeColoursByIndex[$index] ?? null;
    }

    public function setThemeColor(string $colourName, string $colourValue): void
    {
        $colourValue = strtoupper($colourValue);
        $this->themeColours[$colourName] = $colourValue;

        if (isset($this->colourIndexByName[$colourName])) {
            $index = $this->colourIndexByName[$colourName];
            $this->themeColoursByIndex[$index] = $colourValue;

            return;
        }

        $index = count($this->themeColoursByIndex);
        $this->colourIndexByName[$colourName] = $index;
        $this->themeColoursByIndex[$index] = $colourValue;
    }

    public function getMajorFontLatin(): string
    {
        return $this->majorFontLatin;
    }

    public function getMajorFontEastAsian(): string
    {
        return $this->majorFontEastAsian;
    }

    public function getMajorFontComplexScript(): string
    {
        return $this->majorFontComplexScript;
    }

    /**
     * @return array<string, string>
     */
    public function getMajorFontSubstitutions(): array
    {
        return $this->majorFontSubstitutions;
    }

    /**
     * @param array<string, string> $fontSet
     */
    public function setMajorFontValues(string $latinFont, string $eastAsianFont, string $complexScriptFont, array $fontSet): void
    {
        $this->majorFontLatin = $latinFont;
        $this->majorFontEastAsian = $eastAsianFont;
        $this->majorFontComplexScript = $complexScriptFont;
        $this->majorFontSubstitutions = $fontSet;
    }

    public function getMinorFontLatin(): string
    {
        return $this->minorFontLatin;
    }

    public function getMinorFontEastAsian(): string
    {
        return $this->minorFontEastAsian;
    }

    public function getMinorFontComplexScript(): string
    {
        return $this->minorFontComplexScript;
    }

    /**
     * @return array<string, string>
     */
    public function getMinorFontSubstitutions(): array
    {
        return $this->minorFontSubstitutions;
    }

    /**
     * @param array<string, string> $fontSet
     */
    public function setMinorFontValues(string $latinFont, string $eastAsianFont, string $complexScriptFont, array $fontSet): void
    {
        $this->minorFontLatin = $latinFont;
        $this->minorFontEastAsian = $eastAsianFont;
        $this->minorFontComplexScript = $complexScriptFont;
        $this->minorFontSubstitutions = $fontSet;
    }
}
