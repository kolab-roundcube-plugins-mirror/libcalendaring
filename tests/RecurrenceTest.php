<?php

use Roundcube\Tests\StderrMock;

/**
 * libcalendaring_recurrence tests
 *
 * @author Aleksander Machniak <machniak@apheleia-it.ch>
 *
 * Copyright (C) 2022, Apheleia IT AG <contact@apheleia-it.ch>
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU Affero General Public License as
 * published by the Free Software Foundation, either version 3 of the
 * License, or (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 * GNU Affero General Public License for more details.
 *
 * You should have received a copy of the GNU Affero General Public License
 * along with this program. If not, see <http://www.gnu.org/licenses/>.
 */

class RecurrenceTest extends PHPUnit\Framework\TestCase
{
    private $plugin;

    public function setUp(): void
    {
        $rcube = rcmail::get_instance();
        $rcube->plugins->load_plugin('libcalendaring', true, true);

        $this->plugin = $rcube->plugins->get_plugin('libcalendaring');
    }

    /**
     * Data for test_end()
     */
    public function data_end()
    {
        return [
            // non-recurring
            [
                [
                    'recurrence' => [],
                    'start' => new DateTime('2017-08-31 11:00:00'),
                ],
                '2017-08-31 11:00:00', // expected result
            ],
            // daily
            [
                [
                    'recurrence' => ['FREQ' => 'DAILY', 'INTERVAL' => '1', 'COUNT' => 2],
                    'start' => new DateTime('2017-08-31 11:00:00'),
                ],
                '2017-09-01 11:00:00',
            ],
            // weekly
            [
                [
                    'recurrence' => ['FREQ' => 'WEEKLY', 'COUNT' => 3],
                    'start' => new DateTime('2017-08-31 11:00:00'), // Thursday
                ],
                '2017-09-14 11:00:00',
            ],
            // UNTIL
            [
                [
                    'recurrence' => ['FREQ' => 'WEEKLY', 'COUNT' => 3, 'UNTIL' => new DateTime('2017-09-07 11:00:00')],
                    'start' => new DateTime('2017-08-31 11:00:00'), // Thursday
                ],
                '2017-09-07 11:00:00',
            ],
            // Infinite recurrence, no count, no until
            [
                [
                    'recurrence' => ['FREQ' => 'WEEKLY', 'INTERVAL' => '1'],
                    'start' => new DateTime('2017-08-31 11:00:00'), // Thursday
                ],
                '2084-09-21 11:00:00',
            ],

            // TODO: Test an event with EXDATE/RDATEs
        ];
    }
    /**
     * Data for test_first_occurrence()
     */
    public function data_first_occurrence()
    {
        // TODO: BYYEARDAY, BYWEEKNO, BYSETPOS, WKST

        return [
            // non-recurring
            [
                [],                                     // recurrence data
                '2017-08-31 11:00:00',                       // start date
                '2017-08-31 11:00:00',                       // expected result
            ],
            // daily
            [
                ['FREQ' => 'DAILY', 'INTERVAL' => '1'], // recurrence data
                '2017-08-31 11:00:00',                       // start date
                '2017-08-31 11:00:00',                       // expected result
            ],
            // TODO: this one is not supported by the Calendar UI
/*
            array(
                array('FREQ' => 'DAILY', 'INTERVAL' => '1', 'BYMONTH' => 1),
                '2017-08-31 11:00:00',
                '2018-01-01 11:00:00',
            ),
*/
            // weekly
            [
                ['FREQ' => 'WEEKLY', 'INTERVAL' => '1'],
                '2017-08-31 11:00:00', // Thursday
                '2017-08-31 11:00:00',
            ],
            [
                ['FREQ' => 'WEEKLY', 'INTERVAL' => '1', 'BYDAY' => 'WE'],
                '2017-08-31 11:00:00', // Thursday
                '2017-09-06 11:00:00',
            ],
            [
                ['FREQ' => 'WEEKLY', 'INTERVAL' => '1', 'BYDAY' => 'TH'],
                '2017-08-31 11:00:00', // Thursday
                '2017-08-31 11:00:00',
            ],
            [
                ['FREQ' => 'WEEKLY', 'INTERVAL' => '1', 'BYDAY' => 'FR'],
                '2017-08-31 11:00:00', // Thursday
                '2017-09-01 11:00:00',
            ],
            [
                ['FREQ' => 'WEEKLY', 'INTERVAL' => '2'],
                '2017-08-31 11:00:00', // Thursday
                '2017-08-31 11:00:00',
            ],
            [
                ['FREQ' => 'WEEKLY', 'INTERVAL' => '3', 'BYDAY' => 'WE'],
                '2017-08-31 11:00:00', // Thursday
                '2017-09-20 11:00:00',
            ],
            [
                ['FREQ' => 'WEEKLY', 'INTERVAL' => '1', 'BYDAY' => 'WE', 'COUNT' => 1],
                '2017-08-31 11:00:00', // Thursday
                '2017-09-06 11:00:00',
            ],
            [
                ['FREQ' => 'WEEKLY', 'INTERVAL' => '1', 'BYDAY' => 'WE', 'UNTIL' => '2017-09-01'],
                '2017-08-31 11:00:00', // Thursday
                '',
            ],
            // monthly
            [
                ['FREQ' => 'MONTHLY', 'INTERVAL' => '1'],
                '2017-09-08 11:00:00',
                '2017-09-08 11:00:00',
            ],
            [
                ['FREQ' => 'MONTHLY', 'INTERVAL' => '1', 'BYMONTHDAY' => '8,9'],
                '2017-08-31 11:00:00',
                '2017-09-08 11:00:00',
            ],
            [
                ['FREQ' => 'MONTHLY', 'INTERVAL' => '1', 'BYMONTHDAY' => '8,9'],
                '2017-09-08 11:00:00',
                '2017-09-08 11:00:00',
            ],
            [
                ['FREQ' => 'MONTHLY', 'INTERVAL' => '1', 'BYDAY' => '1WE'],
                '2017-08-16 11:00:00',
                '2017-09-06 11:00:00',
            ],
            [
                ['FREQ' => 'MONTHLY', 'INTERVAL' => '1', 'BYDAY' => '-1WE'],
                '2017-08-16 11:00:00',
                '2017-08-30 11:00:00',
            ],
            [
                ['FREQ' => 'MONTHLY', 'INTERVAL' => '2'],
                '2017-09-08 11:00:00',
                '2017-09-08 11:00:00',
            ],
            [
                ['FREQ' => 'MONTHLY', 'INTERVAL' => '2', 'BYMONTHDAY' => '8'],
                '2017-08-31 11:00:00',
                '2017-09-08 11:00:00', // ??????
            ],
            // yearly
            [
                ['FREQ' => 'YEARLY', 'INTERVAL' => '1'],
                '2017-08-16 12:00:00',
                '2017-08-16 12:00:00',
            ],
            [
                ['FREQ' => 'YEARLY', 'INTERVAL' => '1', 'BYMONTH' => '8'],
                '2017-08-16 12:00:00',
                '2017-08-16 12:00:00',
            ],
/*
            // Not supported by Sabre (requires BYMONTH too)
            array(
                array('FREQ' => 'YEARLY', 'INTERVAL' => '1', 'BYDAY' => '-1MO'),
                '2017-08-16 11:00:00',
                '2017-12-25 11:00:00',
            ),
*/
            [
                ['FREQ' => 'YEARLY', 'INTERVAL' => '1', 'BYMONTH' => '8', 'BYDAY' => '-1MO'],
                '2017-08-16 11:00:00',
                '2017-08-28 11:00:00',
            ],
            [
                ['FREQ' => 'YEARLY', 'INTERVAL' => '1', 'BYMONTH' => '1', 'BYDAY' => '1MO'],
                '2017-08-16 11:00:00',
                '2018-01-01 11:00:00',
            ],
            [
                ['FREQ' => 'YEARLY', 'INTERVAL' => '1', 'BYMONTH' => '1,9', 'BYDAY' => '1MO'],
                '2017-08-16 11:00:00',
                '2017-09-04 11:00:00',
            ],
            [
                ['FREQ' => 'YEARLY', 'INTERVAL' => '2'],
                '2017-08-16 11:00:00',
                '2017-08-16 11:00:00',
            ],
            [
                ['FREQ' => 'YEARLY', 'INTERVAL' => '2', 'BYMONTH' => '8'],
                '2017-08-16 11:00:00',
                '2017-08-16 11:00:00',
            ],
/*
            // Not supported by Sabre (requires BYMONTH too)
            array(
                array('FREQ' => 'YEARLY', 'INTERVAL' => '2', 'BYDAY' => '-1MO'),
                '2017-08-16 11:00:00',
                '2017-12-25 11:00:00',
            ),
*/
            // on dates (FIXME: do we really expect the first occurrence to be on the start date?)
            [
                ['RDATE' => [new DateTime('2017-08-10 11:00:00 Europe/Warsaw')]],
                '2017-08-01 11:00:00',
                '2017-08-01 11:00:00',
            ],
        ];
    }

    /**
     * Test for libcalendaring_recurrence::end()
     *
     * @dataProvider data_end
     */
    public function test_end($event, $expected)
    {
        $recurrence = new libcalendaring_recurrence($this->plugin, $event);

        $end = $recurrence->end();

        $this->assertSame($expected, $end ? $end->format('Y-m-d H:i:s') : $end);
    }

    /**
     * Test for libcalendaring_recurrence::first_occurrence()
     *
     * @dataProvider data_first_occurrence
     */
    public function test_first_occurrence($recurrence_data, $start, $expected)
    {
        $start = new DateTime($start);
        if (!empty($recurrence_data['UNTIL'])) {
            $recurrence_data['UNTIL'] = new DateTime($recurrence_data['UNTIL']);
        }

        $recurrence = $this->plugin->get_recurrence();

        StderrMock::start();
        $recurrence->init($recurrence_data, $start);
        $first = $recurrence->first_occurrence();
        StderrMock::stop();

        $this->assertEquals($expected, $first ? $first->format('Y-m-d H:i:s') : '');
    }

    /**
     * Test for libcalendaring_recurrence::first_occurrence() for all-day events
     *
     * @dataProvider data_first_occurrence
     */
    public function test_first_occurrence_allday($recurrence_data, $start, $expected)
    {
        $start = new libcalendaring_datetime($start);
        $start->_dateonly = true;

        if (!empty($recurrence_data['UNTIL'])) {
            $recurrence_data['UNTIL'] = new DateTime($recurrence_data['UNTIL']);
        }

        $recurrence = $this->plugin->get_recurrence();

        StderrMock::start();
        $recurrence->init($recurrence_data, $start);
        $first = $recurrence->first_occurrence();
        StderrMock::stop();

        $this->assertEquals($expected, $first ? $first->format('Y-m-d H:i:s') : '');

        if ($expected) {
            $this->assertTrue($first->_dateonly);
        }
    }

    /**
     * Test for an event with invalid recurrence
     */
    public function test_invalid_recurrence_event()
    {
        date_default_timezone_set('Europe/Berlin');

        // This is an event with no RRULE, but one RDATE, however the RDATE is cancelled by EXDATE.
        // This normally causes Sabre\VObject\Recur\NoInstancesException. We make sure it does not happen.
        // The same will probably happen on any event without recurrence passed to libcalendring_vcalendar.

        $vcal = <<<EOF
            BEGIN:VCALENDAR
            VERSION:2.0
            PRODID:-//Apple Inc.//iCal 5.0.3//EN
            CALSCALE:GREGORIAN
            BEGIN:VEVENT
            UID:fb1cb690-b963-4ea5-b58f-ac9773e36d9a
            DTSTART;TZID=Europe/Berlin:20210604T093000
            DTEND;TZID=Europe/Berlin:20210606T093000
            RDATE:20210604T073000Z
            EXDATE;TZID=Europe/Berlin:20210604T093000
            DTSTAMP:20210528T091628Z
            LAST-MODIFIED:20210528T091628Z
            CREATED:20210528T091213Z
            END:VEVENT
            END:VCALENDAR
            EOF;

        $ical = new libcalendaring_vcalendar();
        $event = $ical->import($vcal)[0];

        $recurrence = new libcalendaring_recurrence($this->plugin, $event);

        $this->assertSame($event['start']->format('Y-m-d H:i:s'), $recurrence->end()->format('Y-m-d H:i:s'));
        $this->assertSame($event['start']->format('Y-m-d H:i:s'), $recurrence->first_occurrence()->format('Y-m-d H:i:s'));
        $this->assertFalse($recurrence->next_start());
        $this->assertFalse($recurrence->next_instance());
    }

    /**
     * Test for libcalendaring_recurrence::next_instance()
     */
    public function test_next_instance()
    {
        date_default_timezone_set('America/New_York');

        $start = new libcalendaring_datetime('2017-08-31 11:00:00', new DateTimeZone('Europe/Berlin'));
        $event = [
            'start'      => $start,
            'recurrence' => ['FREQ' => 'WEEKLY', 'INTERVAL' => '1'],
            'allday'     => true,
        ];

        $recurrence = new libcalendaring_recurrence($this->plugin, $event);
        $next       = $recurrence->next_instance();

        $this->assertEquals($start->format('2017-09-07 H:i:s'), $next['start']->format('Y-m-d H:i:s'), 'Same time');
        $this->assertEquals($start->getTimezone()->getName(), $next['start']->getTimezone()->getName(), 'Same timezone');
        $this->assertTrue($next['start']->_dateonly, '_dateonly flag');
    }

    /**
     * Test for libcalendaring_recurrence::next_instance()
     */
    public function test_next_instance_exdate()
    {
        date_default_timezone_set('America/New_York');

        $start = new libcalendaring_datetime('2023-01-18 10:00:00', new DateTimeZone('Europe/Berlin'));
        $end = new libcalendaring_datetime('2023-01-18 10:30:00', new DateTimeZone('Europe/Berlin'));
        $event = [
            'start' => $start,
            'end' => $end,
            'recurrence' => [
                'FREQ' => 'DAILY',
                'INTERVAL' => '1',
                'EXDATE' => [
                    // Exclude the start date
                    new libcalendaring_datetime('2023-01-18 10:00:00', new DateTimeZone('Europe/Berlin')),
                ],
            ],
        ];

        $recurrence = new libcalendaring_recurrence($this->plugin, $event);

        $next = $recurrence->next_instance();

        $this->assertEquals('2023-01-19 10:00:00', $next['start']->format('Y-m-d H:i:s'));
        $this->assertEquals('Europe/Berlin', $next['start']->getTimezone()->getName());
        $this->assertFalse($next['start']->_dateonly);

        $next = $recurrence->next_instance();

        $this->assertEquals('2023-01-20 10:00:00', $next['start']->format('Y-m-d H:i:s'));
        $this->assertEquals('Europe/Berlin', $next['start']->getTimezone()->getName());
        $this->assertFalse($next['start']->_dateonly);
    }

    /**
     * Test for libcalendaring_recurrence::next_instance()
     */
    public function test_next_instance_dst()
    {
        date_default_timezone_set('America/New_York');

        $start = new libcalendaring_datetime('2021-03-10 10:00:00', new DateTimeZone('Europe/Berlin'));
        $end = new libcalendaring_datetime('2021-03-10 10:30:00', new DateTimeZone('Europe/Berlin'));
        $event = [
            'start' => $start,
            'end' => $end,
            'recurrence' => [
                'FREQ' => 'MONTHLY',
                'INTERVAL' => '1',
            ],
        ];

        $recurrence = new libcalendaring_recurrence($this->plugin, $event);

        $next = $recurrence->next_instance();

        $this->assertEquals('2021-04-10 10:00:00', $next['start']->format('Y-m-d H:i:s'));
        $this->assertEquals('Europe/Berlin', $next['start']->getTimezone()->getName());

        $next = $recurrence->next_instance();

        $this->assertEquals('2021-05-10 10:00:00', $next['start']->format('Y-m-d H:i:s'));
        $this->assertEquals('Europe/Berlin', $next['start']->getTimezone()->getName());

        $start = new libcalendaring_datetime('2021-10-10 10:00:00', new DateTimeZone('Europe/Berlin'));
        $end = new libcalendaring_datetime('2021-10-10 10:30:00', new DateTimeZone('Europe/Berlin'));
        $event = [
            'start' => $start,
            'end' => $end,
            'recurrence' => [
                'FREQ' => 'MONTHLY',
                'INTERVAL' => '1',
            ],
        ];

        $recurrence = new libcalendaring_recurrence($this->plugin, $event);

        $next = $recurrence->next_instance();

        $this->assertEquals('2021-11-10 10:00:00', $next['start']->format('Y-m-d H:i:s'));
        $this->assertEquals('Europe/Berlin', $next['start']->getTimezone()->getName());

        $next = $recurrence->next_instance();

        $this->assertEquals('2021-12-10 10:00:00', $next['start']->format('Y-m-d H:i:s'));
        $this->assertEquals('Europe/Berlin', $next['start']->getTimezone()->getName());

        $next = $recurrence->next_instance();
        $next = $recurrence->next_instance();
        $next = $recurrence->next_instance();
        $next = $recurrence->next_instance();

        $this->assertEquals('2022-04-10 10:00:00', $next['start']->format('Y-m-d H:i:s'));
        $this->assertEquals('Europe/Berlin', $next['start']->getTimezone()->getName());

    }
}
