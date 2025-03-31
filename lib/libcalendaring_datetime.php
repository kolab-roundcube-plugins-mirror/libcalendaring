<?php

/**
 * DateTime wrapper. Main reason for its existence is that
 * you can't set undefined properties on DateTime without
 * a deprecation warning on PHP >= 8.1
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

class libcalendaring_datetime extends DateTime
{
    public $_dateonly = false;

    /**
     * Create an instance from a date string or object
     *
     * @param DateTimeInterface|string $date     Date
     * @param bool                     $dateonly Date only (ignore time)
     */
    public static function createFromAny($date, bool $dateonly = false)
    {
        if (!$date instanceof DateTimeInterface) {
            $date = new DateTime($date, new DateTimeZone('UTC'));
        }

        // Note: On PHP8 we have DateTime::createFromInterface(), but not on PHP7

        $result = self::createFromFormat(
            'Y-m-d\\TH:i:s',
            $date->format('Y-m-d\\TH:i:s'),
            // Sabre will loose timezone on all-day events, use the event start's timezone
            $date->getTimezone()
        );

        $result->_dateonly = $dateonly;

        return $result;
    }
}
