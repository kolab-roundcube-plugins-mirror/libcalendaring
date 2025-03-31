<?php

use Sabre\VObject\Recur\EventIterator;
use Sabre\VObject\Recur\NoInstancesException;

/**
 * Recurrence computation class for shared use.
 *
 * Uitility class to compute recurrence dates from the given rules.
 *
 * @author Aleksander Machniak <machniak@apheleia-it.ch
 *
 * Copyright (C) 2012-2022, Apheleia IT AG <contact@apheleia-it.ch>
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
class libcalendaring_recurrence
{
    protected $lib;
    protected $start;
    protected $engine;
    protected $recurrence;
    protected $dateonly = false;
    protected $event;
    protected $duration;
    protected $isStart = true;

    /**
     * Default constructor
     *
     * @param libcalendaring $lib   The libcalendaring plugin instance
     * @param array          $event The event object to operate on
     */
    public function __construct($lib, $event = null)
    {
        $this->lib   = $lib;
        $this->event = $event;

        if (!empty($event)) {
            if (!empty($event['start']) && is_object($event['start'])
                && !empty($event['end']) && is_object($event['end'])
            ) {
                $this->duration = $event['start']->diff($event['end']);
            }

            $this->init($event['recurrence'], $event['start']);
        }
    }

    /**
     * Initialize recurrence engine
     *
     * @param array    $recurrence The recurrence properties
     * @param DateTime $start      The recurrence start date
     */
    public function init($recurrence, $start)
    {
        $this->start      = $start;
        $this->dateonly   = !empty($start->_dateonly) || !empty($this->event['allday']);
        $this->recurrence = $recurrence;

        $event = [
            'uid' => '1',
            'allday' => $this->dateonly,
            'recurrence' => $recurrence,
            'start' => $start,
            // TODO: END/DURATION ???
            // TODO: moved occurrences ???
        ];

        $vcalendar = new libcalendaring_vcalendar($this->lib->timezone);

        $ve = $vcalendar->toSabreComponent($event);

        try {
            $this->engine = new EventIterator($ve, null, $this->lib->timezone);
        } catch (NoInstancesException $e) {
            // We treat a non-recurring event silently
            // TODO: What about other exceptions?
        }
    }

    /**
     * Get date/time of the next occurence of this event, and push the iterator.
     *
     * @return libcalendaring_datetime|false object or False if recurrence ended
     */
    public function next_start()
    {
        if (empty($this->engine)) {
            return false;
        }

        try {
            $this->engine->next();
            $current = $this->engine->getDtStart();
        } catch (Exception $e) {
            // do nothing
        }

        return !empty($current) ? $this->toDateTime($current) : false;
    }

    /**
     * Get the next recurring instance of this event
     *
     * @return array|false Array with event properties or False if recurrence ended
     */
    public function next_instance()
    {
        if (empty($this->engine)) {
            return false;
        }

        // Here's the workaround for an issue for an event with its start date excluded
        // E.g. A daily event starting on 10th which is one of EXDATE dates
        // should return 11th as next_instance() when called for the first time.
        // Looks like Sabre is setting internal "current date" to 11th on such an object
        // initialization, therefore calling next() would move it to 12th.
        if ($this->isStart && ($next_start = $this->engine->getDtStart())
            && $next_start->format('Ymd') != $this->start->format('Ymd')
        ) {
            $next_start = $this->toDateTime($next_start);
        } else {
            $next_start = $this->next_start();
        }

        $this->isStart = false;

        if ($next_start) {
            $next = $this->event;
            $next['start'] = $next_start;

            if ($this->duration) {
                $next['end'] = clone $next_start;
                $next['end']->add($this->duration);
            }

            $next['recurrence_date'] = clone $next_start;
            $next['_instance'] = libcalendaring::recurrence_instance_identifier($next);

            unset($next['_formatobj']);

            return $next;
        }

        return false;
    }

    /**
     * Get the date of the end of the last occurrence of this recurrence cycle
     *
     * @return libcalendaring_datetime|false End datetime of the last occurrence or False if there's no end date
     */
    public function end()
    {
        if (empty($this->engine)) {
            return $this->toDateTime($this->start);
        }

        // recurrence end date is given
        if (isset($this->recurrence['UNTIL']) && $this->recurrence['UNTIL'] instanceof DateTimeInterface) {
            return $this->toDateTime($this->recurrence['UNTIL']);
        }

        // Run through all items till we reach the end, or limit of iterations
        // Note: Sabre has a limits of iteration in VObject\Settings, so it is not an infinite loop
        try {
            foreach ($this->engine as $end) {
                // do nothing
            }
        } catch (Exception $e) {
            // do nothing
        }

        /*
        if (empty($end) && isset($this->event['start']) && $this->event['start'] instanceof DateTimeInterface) {
            // determine a reasonable end date if none given
            $end = clone $this->event['start'];
            $end->add(new DateInterval('P100Y'));
        }
        */

        return isset($end) ? $this->toDateTime($end) : false;
    }

    /**
     * Find date/time of the first occurrence (excluding start date)
     *
     * @return libcalendaring_datetime|null First occurrence
     */
    public function first_occurrence()
    {
        if (empty($this->engine)) {
            return $this->toDateTime($this->start);
        }

        $start    = clone $this->start;
        $interval = $this->recurrence['INTERVAL'] ?? 1;
        $freq     = $this->recurrence['FREQ'] ?? null;

        switch ($freq) {
            case 'WEEKLY':
                if (empty($this->recurrence['BYDAY'])) {
                    return $start;
                }

                $start->sub(new DateInterval("P{$interval}W"));
                break;

            case 'MONTHLY':
                if (empty($this->recurrence['BYDAY']) && empty($this->recurrence['BYMONTHDAY'])) {
                    return $start;
                }

                $start->sub(new DateInterval("P{$interval}M"));
                break;

            case 'YEARLY':
                if (empty($this->recurrence['BYDAY']) && empty($this->recurrence['BYMONTH'])) {
                    return $start;
                }

                $start->sub(new DateInterval("P{$interval}Y"));
                break;

            case 'DAILY':
                if (!empty($this->recurrence['BYMONTH'])) {
                    break;
                }

                // no break
            default:
                return $start;
        }

        $recurrence = $this->recurrence;

        if (!empty($recurrence['COUNT'])) {
            // Increase count so we do not stop the loop to early
            $recurrence['COUNT'] += 100;
        }

        // Create recurrence that starts in the past
        $self = new self($this->lib);
        $self->init($recurrence, $start);

        // TODO: This method does not work the same way as the kolab_date_recurrence based on
        //       kolabcalendaring. I.e. if an event start date does not match the recurrence rule
        //       it will be returned, kolab_date_recurrence will return the next occurrence in such a case
        //       which is the intended result of this function.
        //       See some commented out test cases in tests/RecurrenceTest.php

        // find the first occurrence
        $found = false;
        while ($next = $self->next_start()) {
            $start = $next;
            if ($next >= $this->start) {
                $found = true;
                break;
            }
        }

        if (!$found) {
            rcube::raise_error(
                [
                    'file' => __FILE__, 'line' => __LINE__,
                    'message' => sprintf(
                        "Failed to find a first occurrence. Start: %s, Recurrence: %s",
                        $this->start->format(DateTime::ISO8601),
                        json_encode($recurrence)
                    ),
                ],
                true
            );

            return null;
        }

        return $this->toDateTime($start);
    }

    /**
     * Convert any DateTime into libcalendaring_datetime
     */
    protected function toDateTime($date, $useStart = true)
    {
        if ($date instanceof DateTimeInterface) {
            $date = libcalendaring_datetime::createFromFormat(
                'Y-m-d\\TH:i:s',
                $date->format('Y-m-d\\TH:i:s'),
                // Sabre will loose timezone on all-day events, use the event start's timezone
                $this->start->getTimezone()
            );
        }

        $date->_dateonly = $this->dateonly;

        if ($useStart && $this->dateonly) {
            // Sabre sets time to 00:00:00 for all-day events,
            // let's copy the time from the event's start
            $date->setTime((int) $this->start->format('H'), (int) $this->start->format('i'), (int) $this->start->format('s'));
        }

        return $date;
    }
}
