<?php

/**
 * Library providing common functions for calendaring plugins
 *
 * Provides utility functions for calendar-related modules such as
 * - alarms display and dismissal
 * - attachment handling
 * - recurrence computation and UI elements
 * - ical parsing and exporting
 * - itip scheduling protocol
 *
 * @version @package_version@
 * @author Thomas Bruederli <bruederli@kolabsys.com>
 *
 * Copyright (C) 2012-2015, Kolab Systems AG <contact@kolabsys.com>
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

class libcalendaring extends rcube_plugin
{
    public $rc;
    public $timezone;
    public $gmt_offset;
    public $dst_active;
    public $timezone_offset;
    public $ical_parts = [];

    /** @var ?rcube_message Email message */
    public $ical_message;

    /** @var array Configuration defaults */
    public $defaults = [
        'calendar_date_format'  => 'Y-m-d',
        'calendar_date_short'   => 'M-j',
        'calendar_date_long'    => 'F j Y',
        'calendar_date_agenda'  => 'l M-d',
        'calendar_time_format'  => 'H:m',
        'calendar_first_day'    => 1,
        'calendar_first_hour'   => 6,
        'calendar_date_format_sets' => [
            'Y-m-d' => ['d M Y', 'm-d', 'l m-d'],
            'Y/m/d' => ['d M Y', 'm/d', 'l m/d'],
            'Y.m.d' => ['d M Y', 'm.d', 'l m.d'],
            'd-m-Y' => ['d M Y', 'd-m', 'l d-m'],
            'd/m/Y' => ['d M Y', 'd/m', 'l d/m'],
            'd.m.Y' => ['d M Y', 'd.m', 'l d.m'],
            'j.n.Y' => ['d M Y', 'd.m', 'l d.m'],
            'm/d/Y' => ['M d Y', 'm/d', 'l m/d'],
        ],
    ];

    private static $instance;

    private $mail_ical_parser;

    /**
     * Singleton getter to allow direct access from other plugins
     */
    public static function get_instance()
    {
        if (!self::$instance) {
            self::$instance = new libcalendaring(rcube::get_instance()->plugins);
            self::$instance->init_instance();
        }

        return self::$instance;
    }

    /**
     * Initializes class properties
     */
    public function init_instance()
    {
        $this->rc = rcube::get_instance();

        // set user's timezone
        try {
            $this->timezone = new DateTimeZone($this->rc->config->get('timezone', 'GMT'));
        } catch (Exception $e) {
            $this->timezone = new DateTimeZone('GMT');
        }

        $now = new DateTime('now', $this->timezone);

        $this->gmt_offset      = $now->getOffset();
        $this->dst_active      = $now->format('I');
        $this->timezone_offset = $this->gmt_offset / 3600 - $this->dst_active;

        $this->add_texts('localization/', false);
    }

    /**
     * Required plugin startup method
     */
    public function init()
    {
        // extend include path to load bundled lib classes
        $include_path = $this->home . '/lib' . PATH_SEPARATOR . ini_get('include_path');
        set_include_path($include_path);

        self::$instance = $this;

        $this->rc = rcube::get_instance();
        $this->init_instance();

        // include client scripts and styles
        if ($this->rc->output) {
            // add hook to display alarms
            $this->add_hook('refresh', [$this, 'refresh']);
            $this->register_action('plugin.alarms', [$this, 'alarms_action']);
            $this->register_action('plugin.expand_attendee_group', [$this, 'expand_attendee_group']);
        }

        // proceed initialization in startup hook
        $this->add_hook('startup', [$this, 'startup']);
    }

    /**
     * Startup hook
     */
    public function startup($args)
    {
        if ($this->rc->output && $this->rc->output->type == 'html') {
            $this->rc->output->set_env('libcal_settings', $this->load_settings());
            $this->include_script('libcalendaring.js');
            $this->include_stylesheet($this->local_skin_path() . '/libcal.css');

            $this->add_label(
                'itipaccepted',
                'itiptentative',
                'itipdeclined',
                'itipdelegated',
                'expandattendeegroup',
                'expandattendeegroupnodata',
                'statusorganizer',
                'statusaccepted',
                'statusdeclined',
                'statusdelegated',
                'statusunknown',
                'statusneeds-action',
                'statustentative',
                'statuscompleted',
                'statusin-process',
                'delegatedto',
                'delegatedfrom',
                'showmore',
                'savein'
            );
        }

        if (($args['task'] ?? null) == 'mail') {
            if ($args['action'] == 'show' || $args['action'] == 'preview') {
                $this->add_hook('message_load', [$this, 'mail_message_load']);
            }
        }
    }

    /**
     * Load iCalendar functions
     *
     * @return libcalendaring_vcalendar iCal parser
     */
    public static function get_ical()
    {
        $self = self::get_instance();
        return new libcalendaring_vcalendar();
    }

    /**
     * Load iTip functions
     */
    public static function get_itip($domain = 'libcalendaring')
    {
        $self = self::get_instance();
        return new libcalendaring_itip($self, $domain);
    }

    /**
     * Load recurrence computation engine
     */
    public static function get_recurrence($object = null)
    {
        $self = self::get_instance();
        return new libcalendaring_recurrence($self, $object);
    }

    /**
     * Shift dates into user's current timezone.
     *
     * @param mixed $dt Any kind of a date representation (DateTime object, string or unix timestamp)
     *
     * @return DateTime DateTime object in user's timezone
     */
    public function adjust_timezone($dt, $dateonly = false)
    {
        if (is_numeric($dt)) {
            $dt = new DateTime('@' . $dt);
        } elseif (is_string($dt)) {
            $dt = rcube_utils::anytodatetime($dt);
        }

        if (($dt instanceof DateTime || $dt instanceof DateTimeImmutable) && empty($dt->_dateonly) && !$dateonly) {
            $dt = $dt->setTimezone($this->timezone);
        }

        return $dt;
    }

    /**
     *
     */
    public function load_settings()
    {
        $this->date_format_defaults();

        $settings = [];
        $keys     = ['date_format', 'time_format', 'date_short', 'date_long', 'date_agenda'];

        foreach ($keys as $key) {
            $settings[$key] = (string)$this->rc->config->get('calendar_' . $key, $this->defaults['calendar_' . $key]);
            $settings[$key] = self::from_php_date_format($settings[$key]);
        }

        $settings['dates_long']  = $settings['date_long'];
        $settings['first_day']   = (int)$this->rc->config->get('calendar_first_day', $this->defaults['calendar_first_day']);
        $settings['timezone']    = $this->timezone_offset;
        $settings['dst']         = $this->dst_active;

        // localization
        $settings['days'] = [
            $this->rc->gettext('sunday'),   $this->rc->gettext('monday'),
            $this->rc->gettext('tuesday'),  $this->rc->gettext('wednesday'),
            $this->rc->gettext('thursday'), $this->rc->gettext('friday'),
            $this->rc->gettext('saturday'),
        ];
        $settings['days_short'] = [
            $this->rc->gettext('sun'), $this->rc->gettext('mon'),
            $this->rc->gettext('tue'), $this->rc->gettext('wed'),
            $this->rc->gettext('thu'), $this->rc->gettext('fri'),
            $this->rc->gettext('sat'),
        ];
        $settings['months'] = [
            $this->rc->gettext('longjan'), $this->rc->gettext('longfeb'),
            $this->rc->gettext('longmar'), $this->rc->gettext('longapr'),
            $this->rc->gettext('longmay'), $this->rc->gettext('longjun'),
            $this->rc->gettext('longjul'), $this->rc->gettext('longaug'),
            $this->rc->gettext('longsep'), $this->rc->gettext('longoct'),
            $this->rc->gettext('longnov'), $this->rc->gettext('longdec'),
        ];
        $settings['months_short'] = [
            $this->rc->gettext('jan'), $this->rc->gettext('feb'),
            $this->rc->gettext('mar'), $this->rc->gettext('apr'),
            $this->rc->gettext('may'), $this->rc->gettext('jun'),
            $this->rc->gettext('jul'), $this->rc->gettext('aug'),
            $this->rc->gettext('sep'), $this->rc->gettext('oct'),
            $this->rc->gettext('nov'), $this->rc->gettext('dec'),
        ];
        $settings['today'] = $this->rc->gettext('today');

        return $settings;
    }


    /**
     * Helper function to set date/time format according to config and user preferences
     */
    private function date_format_defaults()
    {
        static $defaults = [];

        // nothing to be done
        if (isset($defaults['date_format'])) {
            return;
        }

        $defaults['date_format'] = $this->rc->config->get('calendar_date_format', $this->rc->config->get('date_format'));
        $defaults['time_format'] = $this->rc->config->get('calendar_time_format', $this->rc->config->get('time_format'));

        // override defaults
        if ($defaults['date_format']) {
            $this->defaults['calendar_date_format'] = $defaults['date_format'];
        }
        if ($defaults['time_format']) {
            $this->defaults['calendar_time_format'] = $defaults['time_format'];
        }

        // derive format variants from basic date format
        $format_sets = $this->rc->config->get('calendar_date_format_sets', $this->defaults['calendar_date_format_sets']);
        if ($format_set = $format_sets[$this->defaults['calendar_date_format']]) {
            $this->defaults['calendar_date_long'] = $format_set[0];
            $this->defaults['calendar_date_short'] = $format_set[1];
            $this->defaults['calendar_date_agenda'] = $format_set[2];
        }
    }

    /**
     * Compose a date string for the given event
     */
    public function event_date_text($event)
    {
        $fromto  = '--';
        $is_task = !empty($event['_type']) && $event['_type'] == 'task';

        $this->date_format_defaults();

        $date_format = self::to_php_date_format($this->rc->config->get('calendar_date_format', $this->defaults['calendar_date_format']));
        $time_format = self::to_php_date_format($this->rc->config->get('calendar_time_format', $this->defaults['calendar_time_format']));

        $getTimezone = function ($date) {
            if ($newTz = $date->getTimezone()) {
                return $newTz->getName();
            }

            return '';
        };

        $formatDate = function ($date, $format) use ($getTimezone) {
            // This is a workaround for the rcmail::format_date() which does not play nice with timezone
            $tz = $this->rc->config->get('timezone');
            if ($dateTz = $getTimezone($date)) {
                $this->rc->config->set('timezone', $dateTz);
            }
            $result = $this->rc->format_date($date, $format);
            $this->rc->config->set('timezone', $tz);

            return $result;
        };

        // handle task objects
        if ($is_task && !empty($event['due']) && is_object($event['due'])) {
            $fromto = $formatDate($event['due'], !empty($event['due']->_dateonly) ? $date_format : null);

            // add timezone information
            if ($fromto && empty($event['due']->_dateonly) && ($tz = $getTimezone($event['due']))) {
                $fromto .= ' (' . strtr($tz, '_', ' ') . ')';
            }

            return $fromto;
        }

        // abort if no valid event dates are given
        if (!is_object($event['start']) || !is_a($event['start'], 'DateTime') || !is_object($event['end']) || !is_a($event['end'], 'DateTime')) {
            return $fromto;
        }

        if ($event['allday']) {
            $fromto = $formatDate($event['start'], $date_format);
            if (($todate = $formatDate($event['end'], $date_format)) != $fromto) {
                $fromto .= ' - ' . $todate;
            }
        } elseif ($event['start']->format('Ymd') === $event['end']->format('Ymd')) {
            $fromto = $formatDate($event['start'], $date_format) . ' ' . $formatDate($event['start'], $time_format) .
                ' - ' . $formatDate($event['end'], $time_format);
        } else {
            $fromto = $formatDate($event['start'], $date_format) . ' ' . $formatDate($event['start'], $time_format) .
                ' - ' . $formatDate($event['end'], $date_format) . ' ' . $formatDate($event['end'], $time_format);
        }

        // add timezone information
        if ($fromto && empty($event['allday']) && ($tz = $getTimezone($event['start']))) {
            $fromto .= ' (' . strtr($tz, '_', ' ') . ')';
        }

        return $fromto;
    }


    /**
     * Render HTML form for alarm configuration
     */
    public function alarm_select($attrib, $alarm_types, $absolute_time = true)
    {
        unset($attrib['name']);

        $input_value    = new html_inputfield(['name' => 'alarmvalue[]', 'class' => 'edit-alarm-value form-control', 'size' => 3]);
        $input_date     = new html_inputfield(['name' => 'alarmdate[]', 'class' => 'edit-alarm-date form-control', 'size' => 10]);
        $input_time     = new html_inputfield(['name' => 'alarmtime[]', 'class' => 'edit-alarm-time form-control', 'size' => 6]);
        $select_type    = new html_select(['name' => 'alarmtype[]', 'class' => 'edit-alarm-type form-control', 'id' => $attrib['id']]);
        $select_offset  = new html_select(['name' => 'alarmoffset[]', 'class' => 'edit-alarm-offset form-control']);
        $select_related = new html_select(['name' => 'alarmrelated[]', 'class' => 'edit-alarm-related form-control']);
        $object_type    = !empty($attrib['_type']) ? $attrib['_type'] : 'event';

        $select_type->add($this->gettext('none'), '');
        foreach ($alarm_types as $type) {
            $select_type->add($this->gettext(strtolower("alarm{$type}option")), $type);
        }

        foreach (['-M','-H','-D','+M','+H','+D'] as $trigger) {
            $select_offset->add($this->gettext('trigger' . $trigger), $trigger);
        }

        $select_offset->add($this->gettext('trigger0'), '0');
        if ($absolute_time) {
            $select_offset->add($this->gettext('trigger@'), '@');
        }

        $select_related->add($this->gettext('relatedstart'), 'start');
        $select_related->add($this->gettext('relatedend' . $object_type), 'end');

        // pre-set with default values from user settings
        $preset = self::parse_alarm_value($this->rc->config->get('calendar_default_alarm_offset', '-15M'));
        $hidden = ['style' => 'display:none'];

        return html::span(
            'edit-alarm-set',
            $select_type->show($this->rc->config->get('calendar_default_alarm_type', '')) . ' ' .
            html::span(
                ['class' => 'edit-alarm-values input-group', 'style' => 'display:none'],
                $input_value->show($preset[0]) . ' ' .
                $select_offset->show($preset[1]) . ' ' .
                $select_related->show() . ' ' .
                $input_date->show('', $hidden) . ' ' .
                $input_time->show('', $hidden)
            )
        );
    }

    /**
     * Get a list of email addresses of the given user (from login and identities)
     *
     * @param string $user User Email (default to current user)
     *
     * @return array Email addresses related to the user
     */
    public function get_user_emails($user = null)
    {
        static $_emails = [];

        if (empty($user)) {
            $user = $this->rc->user->get_username();
        }

        // return cached result
        if (isset($_emails[$user])) {
            return $_emails[$user];
        }

        $emails = [$user];
        $plugin = $this->rc->plugins->exec_hook('calendar_user_emails', ['emails' => $emails]);
        $emails = array_map('strtolower', $plugin['emails']);

        // add all emails from the current user's identities
        if (!$plugin['abort'] && ($user == $this->rc->user->get_username())) {
            foreach ($this->rc->user->list_emails() as $identity) {
                $emails[] = strtolower($identity['email']);
            }
        }

        $_emails[$user] = array_unique($emails);
        return $_emails[$user];
    }

    /**
     * Set the given participant status to the attendee matching the current user's identities
     * Unsets 'rsvp' flag too.
     *
     * @param array  &$event    Event data
     * @param string $status    The PARTSTAT value to set
     * @param bool   $recursive Recurive call
     *
     * @return string|false Email address of the updated attendee or False if none matching found
     */
    public function set_partstat(&$event, $status, $recursive = true)
    {
        $success = false;
        $emails = $this->get_user_emails();
        foreach ((array)$event['attendees'] as $i => $attendee) {
            if ($attendee['email'] && in_array(strtolower($attendee['email']), $emails)) {
                $event['attendees'][$i]['status'] = strtoupper($status);
                unset($event['attendees'][$i]['rsvp']);
                $success = $attendee['email'];
            }
        }

        // apply partstat update to each existing exception
        if ($event['recurrence'] && is_array($event['recurrence']['EXCEPTIONS'])) {
            foreach ($event['recurrence']['EXCEPTIONS'] as $i => $exception) {
                $this->set_partstat($event['recurrence']['EXCEPTIONS'][$i], $status, false);
            }

            // set link to top-level exceptions
            $event['exceptions'] = &$event['recurrence']['EXCEPTIONS'];
        }

        return $success;
    }


    /*********  Alarms handling  *********/

    /**
     * Helper function to convert alarm trigger strings
     * into two-field values (e.g. "-45M" => 45, "-M")
     */
    public static function parse_alarm_value($val)
    {
        if ($val[0] == '@') {
            return [new DateTime($val)];
        }

        if (preg_match('/([+-]?)P?(T?\d+[HMSDW])+/', $val, $m) && preg_match_all('/T?(\d+)([HMSDW])/', $val, $m2, PREG_SET_ORDER)) {
            if ($m[1] == '') {
                $m[1] = '+';
            }

            $prefix = '';
            foreach ($m2 as $seg) {
                $prefix = $seg[2] == 'D' || $seg[2] == 'W' ? 'P' : 'PT';

                if ($seg[1] > 0) {  // ignore zero values
                    // convert seconds to minutes
                    if ($seg[2] == 'S') {
                        $seg[2] = 'M';
                        $seg[1] = max(1, round(intval($seg[1]) / 60));
                    }

                    return [$seg[1], $m[1] . $seg[2], $m[1] . $seg[1] . $seg[2], $m[1] . $prefix . $seg[1] . $seg[2]];
                }
            }

            // return zero value nevertheless
            return [
                $seg[1] ?? null,
                $m[1] . ($seg[2] ?? ''),
                $m[1] . ($seg[1] ?? '') . ($seg[2] ?? ''),
                $m[1] . $prefix . ($seg[1] ?? '') . ($seg[2] ?? ''),
            ];
        }

        return false;
    }

    /**
     * Convert the alarms list items to be processed on the client
     */
    public static function to_client_alarms($valarms)
    {
        return array_map(function ($alarm) {
            if ($alarm['trigger'] instanceof DateTimeInterface) {
                $alarm['trigger'] = '@' . $alarm['trigger']->format('U');
            } elseif ($trigger = libcalendaring::parse_alarm_value($alarm['trigger'])) {
                $alarm['trigger'] = $trigger[2];
            }
            return $alarm;
        }, (array)$valarms);
    }

    /**
     * Process the alarms values submitted by the client
     */
    public static function from_client_alarms($valarms)
    {
        return array_map(function ($alarm) {
            if ($alarm['trigger'][0] == '@') {
                try {
                    $alarm['trigger'] = new DateTime($alarm['trigger']);
                    $alarm['trigger']->setTimezone(new DateTimeZone('UTC'));
                } catch (Exception $e) { /* handle this ? */
                }
            } elseif ($trigger = libcalendaring::parse_alarm_value($alarm['trigger'])) {
                $alarm['trigger'] = $trigger[3];
            }
            return $alarm;
        }, (array)$valarms);
    }

    /**
     * Render localized text for alarm settings
     */
    public static function alarms_text($alarms)
    {
        if (is_array($alarms) && is_array($alarms[0])) {
            $texts = [];
            foreach ($alarms as $alarm) {
                if ($text = self::alarm_text($alarm)) {
                    $texts[] = $text;
                }
            }

            return implode(', ', $texts);
        } else {
            return self::alarm_text($alarms);
        }
    }

    /**
     * Render localized text for a single alarm property
     */
    public static function alarm_text($alarm)
    {
        $related = null;

        if (is_string($alarm)) {
            [$trigger, $action] = explode(':', $alarm);
        } else {
            $trigger = $alarm['trigger'];
            $action  = $alarm['action'];

            if (!empty($alarm['related'])) {
                $related = $alarm['related'];
            }
        }

        $text  = '';
        $rcube = rcmail::get_instance();

        switch ($action) {
            case 'EMAIL':
                $text = $rcube->gettext('libcalendaring.alarmemail');
                break;
            case 'DISPLAY':
                $text = $rcube->gettext('libcalendaring.alarmdisplay');
                break;
            case 'AUDIO':
                $text = $rcube->gettext('libcalendaring.alarmaudio');
                break;
        }

        if ($trigger instanceof DateTimeInterface) {
            $text .= ' ' . $rcube->gettext([
                'name' => 'libcalendaring.alarmat',
                'vars' => ['datetime' => $rcube->format_date($trigger)],
            ]);
        } elseif (preg_match('/@(\d+)/', $trigger, $m)) {
            $text .= ' ' . $rcube->gettext([
                'name' => 'libcalendaring.alarmat',
                'vars' => ['datetime' => $rcube->format_date($m[1])],
            ]);
        } elseif ($val = self::parse_alarm_value($trigger)) {
            $r = $related && strtoupper($related) == 'END' ? 'end' : '';
            // TODO: for all-day events say 'on date of event at XX' ?
            if ($val[0] == 0) {
                $text .= ' ' . $rcube->gettext('libcalendaring.triggerattime' . $r);
            } else {
                $label = 'libcalendaring.trigger' . $r . $val[1];
                $text .= ' ' . intval($val[0]) . ' ' . $rcube->gettext($label);
            }
        } else {
            return false;
        }

        return $text;
    }

    /**
     * Get the next alarm (time & action) for the given event
     *
     * @param array $rec Record data
     *
     * @return array|null Hash array with alarm time/type or null if no alarms are configured
     */
    public static function get_next_alarm($rec, $type = 'event')
    {
        if (
            (empty($rec['valarms']) && empty($rec['alarms']))
            || !empty($rec['cancelled'])
            || (!empty($rec['status']) && $rec['status'] == 'CANCELLED')
        ) {
            return null;
        }

        if ($type == 'task') {
            $timezone = self::get_instance()->timezone;
            if (!empty($rec['startdate'])) {
                $time = !empty($rec['starttime']) ? $rec['starttime'] : '12:00';
                $rec['start'] = new DateTime($rec['startdate'] . ' ' . $time, $timezone);
            }
            if (!empty($rec['date'])) {
                $time = !empty($rec['time']) ? $rec['time'] : '12:00';
                $rec[!empty($rec['start']) ? 'end' : 'start'] = new DateTime($rec['date'] . ' ' . $time, $timezone);
            }
        }

        if (empty($rec['end'])) {
            $rec['end'] = $rec['start'];
        }

        // support legacy format
        if (empty($rec['valarms'])) {
            [$trigger, $action] = explode(':', $rec['alarms'], 2);
            if ($alarm = self::parse_alarm_value($trigger)) {
                $rec['valarms'] = [['action' => $action, 'trigger' => $alarm[3] ?: $alarm[0]]];
            }
        }

        $expires = new DateTime('now - 12 hours');
        $result = null;
        $notify_at = null;

        // handle multiple alarms
        foreach ($rec['valarms'] as $alarm) {
            $notify_time = null;

            if ($alarm['trigger'] instanceof DateTimeInterface) {
                $notify_time = $alarm['trigger'];
            } elseif (is_string($alarm['trigger'])) {
                $refdate = !empty($alarm['related']) && $alarm['related'] == 'END' ? $rec['end'] : $rec['start'];

                // abort if no reference date is available to compute notification time
                if (!is_a($refdate, 'DateTime')) {
                    continue;
                }

                // TODO: for all-day events, take start @ 00:00 as reference date ?

                try {
                    $interval = new DateInterval(trim($alarm['trigger'], '+-'));
                    $interval->invert = $alarm['trigger'][0] == '-' ? 1 : 0;
                    $notify_time = clone $refdate;
                    $notify_time->add($interval);
                } catch (Exception $e) {
                    rcube::raise_error($e, true);
                    continue;
                }
            }

            if ($notify_time && (!$notify_at || ($notify_time > $notify_at && $notify_time > $expires))) {
                $alarm_id = substr(md5($rec['id'] ?? ($rec['uid'] ?? 'none')), 0, 16);
                $notify_at = $notify_time;

                // generate a unique alarm ID if multiple alarms are set
                if (count($rec['valarms']) > 1) {
                    $alarm_id = $alarm_id . '-' . $notify_at->format('Ymd\THis');
                }

                $result = [
                    'time'   => $notify_at->format('U'),
                    'action' => !empty($alarm['action']) ? strtoupper($alarm['action']) : 'DISPLAY',
                    'id'     => $alarm_id,
                    'prop'   => $alarm,
                ];
            }
        }

        return $result;
    }

    /**
     * Handler for keep-alive requests
     * This will check for pending notifications and pass them to the client
     */
    public function refresh($attr)
    {
        // collect pending alarms from all providers (e.g. calendar, tasks)
        $plugin = $this->rc->plugins->exec_hook('pending_alarms', [
            'time' => time(),
            'alarms' => [],
        ]);

        if (!$plugin['abort'] && !empty($plugin['alarms'])) {
            // make sure texts and env vars are available on client
            $this->add_texts('localization/', true);
            $this->rc->output->add_label('close');
            $this->rc->output->set_env('snooze_select', $this->snooze_select());
            $this->rc->output->command('plugin.display_alarms', $this->_alarms_output($plugin['alarms']));
        }
    }

    /**
     * Handler for alarm dismiss/snooze requests
     */
    public function alarms_action()
    {
        $data = rcube_utils::get_input_value('data', rcube_utils::INPUT_POST, true);

        $data['ids'] = explode(',', $data['id']);
        $data['success'] = false;

        $plugin = $this->rc->plugins->exec_hook('dismiss_alarms', $data);

        if (!empty($plugin['success'])) {
            $this->rc->output->show_message('successfullysaved', 'confirmation');
        } else {
            $this->rc->output->show_message('calendar.errorsaving', 'error');
        }
    }

    /**
     * Generate reduced and streamlined output for pending alarms
     */
    private function _alarms_output($alarms)
    {
        $out = [];
        foreach ($alarms as $alarm) {
            $out[] = [
                'id'       => $alarm['id'],
                'start'    => !empty($alarm['start']) ? $this->adjust_timezone($alarm['start'])->format('c') : '',
                'end'      => !empty($alarm['end']) ? $this->adjust_timezone($alarm['end'])->format('c') : '',
                'allDay'   => !empty($alarm['allday']),
                'action'   => $alarm['action'],
                'title'    => $alarm['title'],
                'location' => $alarm['location'] ?? null,
            ];
        }

        return $out;
    }

    /**
     * Render a dropdown menu to choose snooze time
     */
    private function snooze_select($attrib = [])
    {
        $steps = [
             5 => 'repeatinmin',
            10 => 'repeatinmin',
            15 => 'repeatinmin',
            20 => 'repeatinmin',
            30 => 'repeatinmin',
            60 => 'repeatinhr',
            120 => 'repeatinhrs',
            1440 => 'repeattomorrow',
            10080 => 'repeatinweek',
        ];

        $items = [];
        foreach ($steps as $n => $label) {
            $items[] = html::tag('li', null, html::a(
                ['href' => "#" . ($n * 60), 'class' => 'active'],
                $this->gettext(['name' => $label, 'vars' => ['min' => $n % 60, 'hrs' => intval($n / 60)]])
            ));
        }

        return html::tag('ul', $attrib + ['class' => 'toolbarmenu menu'], implode("\n", $items), html::$common_attrib);
    }


    /*********  Recurrence rules handling ********/

    /**
     * Render localized text describing the recurrence rule of an event
     */
    public function recurrence_text($rrule)
    {
        $limit     = 10;
        $exdates   = [];
        $format    = $this->rc->config->get('calendar_date_format', $this->defaults['calendar_date_format']);
        $format    = self::to_php_date_format($format);
        $format_fn = function ($dt) use ($format) {
            return rcmail::get_instance()->format_date($dt, $format);
        };

        if (!empty($rrule['EXDATE']) && is_array($rrule['EXDATE'])) {
            $exdates = array_map($format_fn, $rrule['EXDATE']);
        }

        if (empty($rrule['FREQ']) && !empty($rrule['RDATE'])) {
            $rdates = array_map($format_fn, $rrule['RDATE']);
            $more   = false;

            if (!empty($exdates)) {
                $rdates = array_diff($rdates, $exdates);
            }

            if (count($rdates) > $limit) {
                $rdates = array_slice($rdates, 0, $limit);
                $more   = true;
            }

            return $this->gettext('ondate') . ' ' . implode(', ', $rdates) . ($more ? '...' : '');
        }

        $output  = sprintf('%s %d ', $this->gettext('every'), !empty($rrule['INTERVAL']) ? $rrule['INTERVAL'] : 1);

        switch ($rrule['FREQ']) {
            case 'DAILY':
                $output .= $this->gettext('days');
                break;
            case 'WEEKLY':
                $output .= $this->gettext('weeks');
                break;
            case 'MONTHLY':
                $output .= $this->gettext('months');
                break;
            case 'YEARLY':
                $output .= $this->gettext('years');
                break;
        }

        if (!empty($rrule['COUNT'])) {
            $until = $this->gettext(['name' => 'forntimes', 'vars' => ['nr' => $rrule['COUNT']]]);
        } elseif (!empty($rrule['UNTIL'])) {
            $until = $this->gettext('recurrencend') . ' ' . $this->rc->format_date($rrule['UNTIL'], $format);
        } else {
            $until = $this->gettext('forever');
        }

        $output .= ', ' . $until;

        if (!empty($exdates)) {
            $more = false;
            if (count($exdates) > $limit) {
                $exdates = array_slice($exdates, 0, $limit);
                $more    = true;
            }

            $output  .= '; ' . $this->gettext('except') . ' ' . implode(', ', $exdates) . ($more ? '...' : '');
        }

        return $output;
    }

    /**
     * Generate the form for recurrence settings
     */
    public function recurrence_form($attrib = [])
    {
        $html = '';
        switch ($attrib['part']) {
            // frequency selector
            case 'frequency':
                $select = new html_select(['name' => 'frequency', 'id' => 'edit-recurrence-frequency', 'class' => 'form-control']);
                $select->add($this->gettext('never'), '');
                $select->add($this->gettext('daily'), 'DAILY');
                $select->add($this->gettext('weekly'), 'WEEKLY');
                $select->add($this->gettext('monthly'), 'MONTHLY');
                $select->add($this->gettext('yearly'), 'YEARLY');
                $select->add($this->gettext('rdate'), 'RDATE');
                $html = html::label(['for' => 'edit-recurrence-frequency', 'class' => 'col-form-label col-sm-2'], $this->gettext('frequency'))
                    . html::div('col-sm-10', $select->show(''));
                break;

                // daily recurrence
            case 'daily':
                $select = $this->interval_selector(['name' => 'interval', 'class' => 'edit-recurrence-interval form-control', 'id' => 'edit-recurrence-interval-daily']);
                $html = html::div($attrib, html::label(['for' => 'edit-recurrence-interval-daily', 'class' => 'col-form-label col-sm-2'], $this->gettext('every'))
                    . html::div('col-sm-10 input-group', $select->show(1) . html::span('label-after input-group-append', html::span('input-group-text', $this->gettext('days')))));
                break;

                // weekly recurrence form
            case 'weekly':
                $select = $this->interval_selector(['name' => 'interval', 'class' => 'edit-recurrence-interval form-control', 'id' => 'edit-recurrence-interval-weekly']);
                $html = html::div($attrib, html::label(['for' => 'edit-recurrence-interval-weekly', 'class' => 'col-form-label col-sm-2'], $this->gettext('every'))
                    . html::div('col-sm-10 input-group', $select->show(1) . html::span('label-after input-group-append', html::span('input-group-text', $this->gettext('weeks')))));

                // weekday selection
                $daymap   = ['sun','mon','tue','wed','thu','fri','sat'];
                $checkbox = new html_checkbox(['name' => 'byday', 'class' => 'edit-recurrence-weekly-byday']);
                $first    = $this->rc->config->get('calendar_first_day', 1);

                for ($weekdays = '', $j = $first; $j <= $first + 6; $j++) {
                    $d = $j % 7;
                    $weekdays .= html::label(
                        ['class' => 'weekday'],
                        $checkbox->show('', ['value' => strtoupper(substr($daymap[$d], 0, 2))]) .
                        $this->gettext($daymap[$d])
                    ) . ' ';
                }

                $html .= html::div($attrib, html::label(['class' => 'col-form-label col-sm-2'], $this->gettext('bydays'))
                    . html::div('col-sm-10 form-control-plaintext', $weekdays));
                break;

                // monthly recurrence form
            case 'monthly':
                $select = $this->interval_selector(['name' => 'interval', 'class' => 'edit-recurrence-interval form-control', 'id' => 'edit-recurrence-interval-monthly']);
                $html = html::div($attrib, html::label(['for' => 'edit-recurrence-interval-monthly', 'class' => 'col-form-label col-sm-2'], $this->gettext('every'))
                    . html::div('col-sm-10 input-group', $select->show(1) . html::span('label-after input-group-append', html::span('input-group-text', $this->gettext('months')))));

                $checkbox = new html_checkbox(['name' => 'bymonthday', 'class' => 'edit-recurrence-monthly-bymonthday']);
                for ($monthdays = '', $d = 1; $d <= 31; $d++) {
                    $monthdays .= html::label(['class' => 'monthday'], $checkbox->show('', ['value' => $d]) . $d);
                    $monthdays .= $d % 7 ? ' ' : html::br();
                }

                // rule selectors
                $radio = new html_radiobutton(['name' => 'repeatmode', 'class' => 'edit-recurrence-monthly-mode']);
                $table = new html_table(['cols' => 2, 'border' => 0, 'cellpadding' => 0, 'class' => 'formtable']);
                $table->add('label', html::label(null, $radio->show('BYMONTHDAY', ['value' => 'BYMONTHDAY']) . ' ' . $this->gettext('each')));
                $table->add(null, $monthdays);
                $table->add('label', html::label(null, $radio->show('', ['value' => 'BYDAY']) . ' ' . $this->gettext('every')));
                $table->add('recurrence-onevery', $this->rrule_selectors($attrib['part']));

                $html .= html::div($attrib, html::label(['class' => 'col-form-label col-sm-2'], $this->gettext('bydays'))
                    . html::div('col-sm-10 form-control-plaintext', $table->show()));
                break;

                // annually recurrence form
            case 'yearly':
                $select = $this->interval_selector(['name' => 'interval', 'class' => 'edit-recurrence-interval form-control', 'id' => 'edit-recurrence-interval-yearly']);
                $html = html::div($attrib, html::label(['for' => 'edit-recurrence-interval-yearly', 'class' => 'col-form-label col-sm-2'], $this->gettext('every'))
                    . html::div('col-sm-10 input-group', $select->show(1) . html::span('label-after input-group-append', html::span('input-group-text', $this->gettext('years')))));

                // month selector
                $monthmap = ['','jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'];
                $checkbox = new html_checkbox(['name' => 'bymonth', 'class' => 'edit-recurrence-yearly-bymonth']);

                for ($months = '', $m = 1; $m <= 12; $m++) {
                    $months .= html::label(['class' => 'month'], $checkbox->show(null, ['value' => $m]) . $this->gettext($monthmap[$m]));
                    $months .= $m % 4 ? ' ' : html::br();
                }

                $html .= html::div($attrib, html::label(['class' => 'col-form-label col-sm-2'], $this->gettext('bymonths'))
                    . html::div(
                        'col-sm-10 form-control-plaintext',
                        html::div(['id' => 'edit-recurrence-yearly-bymonthblock'], $months)
                        . html::div('recurrence-onevery', $this->rrule_selectors($attrib['part'], '---'))
                    ));
                break;

                // end of recurrence form
            case 'until':
                $radio  = new html_radiobutton(['name' => 'repeat', 'class' => 'edit-recurrence-until']);
                $select = $this->interval_selector(['name' => 'times', 'id' => 'edit-recurrence-repeat-times', 'class' => 'form-control']);
                $input  = new html_inputfield(['name' => 'untildate', 'id' => 'edit-recurrence-enddate', 'size' => '10', 'class' => 'form-control datepicker']);

                $html = html::div(
                    'line first',
                    $radio->show('', ['value' => '', 'id' => 'edit-recurrence-repeat-forever'])
                        . ' ' . html::label('edit-recurrence-repeat-forever', $this->gettext('forever'))
                );

                $label = $this->gettext('ntimes');
                if (strpos($label, '$') === 0) {
                    $label = str_replace('$n', '', $label);
                    $group  = $select->show(1)
                        . html::span('input-group-append', html::span('input-group-text', rcube::Q($label)));
                } else {
                    $label = str_replace('$n', '', $label);
                    $group  = html::span('input-group-prepend', html::span('input-group-text', rcube::Q($label)))
                        . $select->show(1);
                }

                $html .= html::div(
                    'line',
                    $radio->show('', ['value' => 'count', 'id' => 'edit-recurrence-repeat-count'])
                        . ' ' . html::label('edit-recurrence-repeat-count', $this->gettext('for'))
                        . ' ' . html::span('input-group', $group)
                );

                $html .= html::div(
                    'line',
                    $radio->show('', ['value' => 'until', 'id' => 'edit-recurrence-repeat-until', 'aria-label' => $this->gettext('untilenddate')])
                        . ' ' . html::label('edit-recurrence-repeat-until', $this->gettext('untildate'))
                        . ' ' . $input->show('', ['aria-label' => $this->gettext('untilenddate')])
                );

                $html = html::div($attrib, html::label(['class' => 'col-form-label col-sm-2'], ucfirst($this->gettext('recurrencend')))
                    . html::div('col-sm-10', $html));
                break;

            case 'rdate':
                $ul     = html::tag('ul', ['id' => 'edit-recurrence-rdates', 'class' => 'recurrence-rdates'], '');
                $input  = new html_inputfield(['name' => 'rdate', 'id' => 'edit-recurrence-rdate-input', 'size' => "10", 'class' => 'form-control datepicker']);
                $button = new html_inputfield(['type' => 'button', 'class' => 'button add', 'value' => $this->gettext('addrdate')]);

                $html = html::div($attrib, html::label(['class' => 'col-form-label col-sm-2', 'for' => 'edit-recurrence-rdate-input'], $this->gettext('bydates'))
                    . html::div('col-sm-10', $ul . html::div('inputform', $input->show() . $button->show())));
                break;
        }

        return $html;
    }

    /**
     * Input field for interval selection
     */
    private function interval_selector($attrib)
    {
        $select = new html_select($attrib);
        $select->add(range(1, 30), range(1, 30));
        return $select;
    }

    /**
     * Drop-down menus for recurrence rules like "each last sunday of"
     */
    private function rrule_selectors($part, $noselect = null)
    {
        // rule selectors
        $select_prefix = new html_select(['name' => 'bydayprefix', 'id' => "edit-recurrence-$part-prefix", 'class' => 'form-control']);
        if ($noselect) {
            $select_prefix->add($noselect, '');
        }
        $select_prefix->add(
            [
                $this->gettext('first'),
                $this->gettext('second'),
                $this->gettext('third'),
                $this->gettext('fourth'),
                $this->gettext('last'),
            ],
            [1, 2, 3, 4, -1]
        );

        $select_wday = new html_select(['name' => 'byday', 'id' => "edit-recurrence-$part-byday", 'class' => 'form-control']);
        if ($noselect) {
            $select_wday->add($noselect, '');
        }

        $daymap = ['sunday','monday','tuesday','wednesday','thursday','friday','saturday'];
        $first = $this->rc->config->get('calendar_first_day', 1);
        for ($j = $first; $j <= $first + 6; $j++) {
            $d = $j % 7;
            $select_wday->add($this->gettext($daymap[$d]), strtoupper(substr($daymap[$d], 0, 2)));
        }

        return $select_prefix->show() . '&nbsp;' . $select_wday->show();
    }

    /**
     * Convert the recurrence settings to be processed on the client
     */
    public function to_client_recurrence($recurrence, $allday = false)
    {
        if (!empty($recurrence['UNTIL'])) {
            $recurrence['UNTIL'] = $this->adjust_timezone($recurrence['UNTIL'], $allday)->format('c');
        }

        // format RDATE values
        if (!empty($recurrence['RDATE'])) {
            $libcal = $this;
            $recurrence['RDATE'] = array_map(function ($rdate) use ($libcal) {
                return $libcal->adjust_timezone($rdate, true)->format('c');
            }, (array) $recurrence['RDATE']);
        }

        unset($recurrence['EXCEPTIONS']);

        return $recurrence;
    }

    /**
     * Process the alarms values submitted by the client
     */
    public function from_client_recurrence($recurrence, $start = null)
    {
        if (is_array($recurrence) && !empty($recurrence['UNTIL'])) {
            $recurrence['UNTIL'] = new DateTime($recurrence['UNTIL'], $this->timezone);
        }

        if (is_array($recurrence) && !empty($recurrence['RDATE'])) {
            $tz = $this->timezone;
            $recurrence['RDATE'] = array_map(function ($rdate) use ($tz, $start) {
                try {
                    $dt = new DateTime($rdate, $tz);
                    if (is_a($start, 'DateTime')) {
                        $dt->setTime($start->format('G'), $start->format('i'));
                    }
                    return $dt;
                } catch (Exception $e) {
                    return null;
                }
            }, $recurrence['RDATE']);
        }

        return $recurrence;
    }


    /*********  iTip message detection  *********/

    /**
     * Check mail message structure of there are .ics files attached
     */
    public function mail_message_load($p)
    {
        $this->ical_parts = [];
        $this->mail_ical_parser = null;
        $this->ical_message = $p['object'];
        $itip_part          = null;

        // check all message parts for .ics files
        foreach ((array)$this->ical_message->mime_parts as $part) {
            /** @var rcube_message_part $part */
            if (self::part_is_vcalendar($part, $this->ical_message)) {
                if (!empty($part->ctype_parameters['method'])) {
                    $itip_part = $part->mime_id;
                } else {
                    $this->ical_parts[] = $part->mime_id;
                }
            }
        }

        // priorize part with method parameter
        if ($itip_part) {
            $this->ical_parts = [$itip_part];
        }
    }

    /**
     * Getter for the parsed iCal objects attached to the current email message
     *
     * @return libcalendaring_vcalendar Parser instance with the parsed objects
     */
    public function get_mail_ical_objects()
    {
        // create parser and load ical objects
        if (!$this->mail_ical_parser) {
            $this->mail_ical_parser = $this->get_ical();

            foreach ($this->ical_parts as $mime_id) {
                $part    = $this->ical_message->mime_parts[$mime_id];
                $charset = ($part->ctype_parameters['charset'] ?? '') ?: RCUBE_CHARSET;
                $body = $this->ical_message->get_part_body($mime_id, true);

                if ($body === null || $body === false) {
                    rcube::raise_error("Failed to get (iTip) body for message part: {$this->ical_message->uid}/{$mime_id}", true);
                    continue;
                }

                $this->mail_ical_parser->import($body, $charset);

                // check if the parsed object is an instance of a recurring event/task
                array_walk($this->mail_ical_parser->objects, 'libcalendaring::identify_recurrence_instance');

                // stop on the part that has an iTip method specified
                if (count($this->mail_ical_parser->objects) && $this->mail_ical_parser->method) {
                    $this->mail_ical_parser->message_date = $this->ical_message->headers->date;
                    $this->mail_ical_parser->mime_id = $mime_id;

                    // store the message's sender address for comparisons
                    $from = rcube_mime::decode_address_list($this->ical_message->headers->from, 1, true, null, true);
                    $this->mail_ical_parser->sender = !empty($from) ? $from[1] : '';

                    if (!empty($this->mail_ical_parser->sender)) {
                        foreach ($this->mail_ical_parser->objects as $i => $object) {
                            $this->mail_ical_parser->objects[$i]['_sender'] = $this->mail_ical_parser->sender;
                            $this->mail_ical_parser->objects[$i]['_sender_utf'] = rcube_utils::idn_to_utf8($this->mail_ical_parser->sender);
                        }
                    }

                    break;
                }
            }
        }

        return $this->mail_ical_parser;
    }

    /**
     * Read the given mime message from IMAP and parse ical data
     *
     * @param string $mbox    Mailbox name
     * @param string $uid     Message UID
     * @param string $mime_id Message part ID and object index (e.g. '1.2:0')
     * @param string $type    Object type filter (optional)
     *
     * @return ?array Hash array with the parsed iCal
     */
    public function mail_get_itip_object($mbox, $uid, $mime_id, $type = null)
    {
        if (empty($uid) || empty($mime_id)) {
            return null;
        }

        $charset = RCUBE_CHARSET;

        // establish imap connection
        $imap = $this->rc->get_storage();
        $imap->set_folder($mbox);

        [$mime_id, $index] = explode(':', $mime_id);

        $part    = $imap->get_message_part($uid, $mime_id);
        $headers = $imap->get_message_headers($uid);
        $parser  = $this->get_ical();

        if ($part) {
            if (!empty($part->ctype_parameters['charset'])) {
                $charset = $part->ctype_parameters['charset'];
            }

            $objects = $parser->import($part, $charset);
        }

        // successfully parsed events/tasks?
        if (!empty($objects) && ($object = $objects[$index]) && (!$type || $object['_type'] == $type)) {
            if ($parser->method) {
                $object['_method'] = $parser->method;
            }

            // store the message's sender address for comparisons
            $from = rcube_mime::decode_address_list($headers->from, 1, true, null, true);
            $object['_sender'] = !empty($from) ? $from[1] : '';
            $object['_sender_utf'] = rcube_utils::idn_to_utf8($object['_sender']);

            // check if this is an instance of a recurring event/task
            self::identify_recurrence_instance($object);

            return $object;
        }

        return null;
    }

    /**
     * Checks if specified message part is a vcalendar data
     *
     * @param rcube_message_part $part    Part object
     * @param rcube_message      $message Message object
     *
     * @return bool True if part is of type vcard
     */
    public static function part_is_vcalendar($part, $message = null)
    {
        // First check if the message is "valid" (i.e. not multipart/report)
        if ($message) {
            $level = explode('.', $part->mime_id);

            while (array_pop($level) !== null) {
                $id     = implode('.', $level) ?: 0;
                $parent = !empty($message->mime_parts[$id]) ? $message->mime_parts[$id] : null;
                if ($parent && $parent->mimetype == 'multipart/report') {
                    return false;
                }
            }
        }

        return (
            in_array($part->mimetype, ['text/calendar', 'text/x-vcalendar', 'application/ics']) ||
            // Apple sends files as application/x-any (!?)
            ($part->mimetype == 'application/x-any' && !empty($part->filename) && preg_match('/\.ics$/i', $part->filename))
        );
    }

    /**
     * Single occourrences of recurring events are identified by their RECURRENCE-ID property
     * in iCal which is represented as 'recurrence_date' in our internal data structure.
     *
     * Check if such a property exists and derive the '_instance' identifier and '_savemode'
     * attributes which are used in the storage backend to identify the nested exception item.
     */
    public static function identify_recurrence_instance(&$object)
    {
        // for savemode=all, remove recurrence instance identifiers
        if (!empty($object['_savemode']) && $object['_savemode'] == 'all' && !empty($object['recurrence'])) {
            unset($object['_instance'], $object['recurrence_date']);
        }
        // set instance and 'savemode' according to recurrence-id
        elseif (!empty($object['recurrence_date']) && $object['recurrence_date'] instanceof DateTimeInterface) {
            $object['_instance'] = self::recurrence_instance_identifier($object);
            $object['_savemode'] = !empty($object['thisandfuture']) ? 'future' : 'current';
        } elseif (!empty($object['recurrence_id']) && !empty($object['_instance'])) {
            if (strlen($object['_instance']) > 4) {
                $object['recurrence_date'] = rcube_utils::anytodatetime($object['_instance'], $object['start']->getTimezone());
            } else {
                $object['recurrence_date'] = clone $object['start'];
            }
        }
    }

    /**
     * Return a date() format string to render identifiers for recurrence instances
     *
     * @param array $event Hash array with event properties
     *
     * @return string Format string
     */
    public static function recurrence_id_format($event)
    {
        return !empty($event['allday']) ? 'Ymd' : 'Ymd\THis';
    }

    /**
     * Return the identifer for the given instance of a recurring event
     *
     * @param array $event  Hash array with event properties
     * @param ?bool $allday All-day flag from the main event
     *
     * @return mixed Format string or null if identifier cannot be generated
     */
    public static function recurrence_instance_identifier($event, $allday = null)
    {
        $instance_date = !empty($event['recurrence_date']) ? $event['recurrence_date'] : $event['start'];

        if ($instance_date instanceof DateTimeInterface) {
            // According to RFC5545 (3.8.4.4) RECURRENCE-ID format should
            // be date/date-time depending on the main event type, not the exception
            if ($allday === null) {
                $allday = !empty($event['allday']);
            }

            return $instance_date->format($allday ? 'Ymd' : 'Ymd\THis');
        }
    }

    /**
     * Check if a specified event is "identical" to the specified recurrence exception
     *
     * @param array $event     Hash array with occurrence properties
     * @param array $exception Hash array with exception properties
     *
     * @return bool
     */
    public static function is_recurrence_exception($event, $exception)
    {
        $instance_date = !empty($event['recurrence_date']) ? $event['recurrence_date'] : $event['start'];
        $exception_date = !empty($exception['recurrence_date']) ? $exception['recurrence_date'] : $exception['start'];

        if ($instance_date instanceof DateTimeInterface && $exception_date instanceof DateTimeInterface) {
            // Timezone???
            return $instance_date->format('Ymd') === $exception_date->format('Ymd');
        }

        return false;
    }


    /*********  Attendee handling functions  *********/

    /**
     * Handler for attendee group expansion requests
     */
    public function expand_attendee_group()
    {
        $id     = rcube_utils::get_input_value('id', rcube_utils::INPUT_POST);
        $data   = rcube_utils::get_input_value('data', rcube_utils::INPUT_POST, true);
        $result = ['id' => $id, 'members' => []];
        $maxnum = 500;

        // iterate over all autocomplete address books (we don't know the source of the group)
        foreach ((array)$this->rc->config->get('autocomplete_addressbooks', 'sql') as $abook_id) {
            if (($abook = $this->rc->get_address_book($abook_id)) && $abook->groups) {
                foreach ($abook->list_groups($data['name'], 1) as $group) {
                    // this is the matching group to expand
                    if (in_array($data['email'], (array)$group['email'])) {
                        $abook->set_pagesize($maxnum);
                        $abook->set_group($group['ID']);

                        // get all members
                        $res = $abook->list_records($this->rc->config->get('contactlist_fields'));

                        // handle errors (e.g. sizelimit, timelimit)
                        if ($abook->get_error()) {
                            $result['error'] = $this->rc->gettext('expandattendeegrouperror', 'libcalendaring');
                            $res = false;
                        }
                        // check for maximum number of members (we don't wanna bloat the UI too much)
                        elseif ($res->count > $maxnum) {
                            $result['error'] = $this->rc->gettext('expandattendeegroupsizelimit', 'libcalendaring');
                            $res = false;
                        }

                        while ($res && ($member = $res->iterate())) {
                            $emails = (array)$abook->get_col_values('email', $member, true);
                            if (!empty($emails) && ($email = array_shift($emails))) {
                                $result['members'][] = [
                                    'email' => $email,
                                    'name' => rcube_addressbook::compose_list_name($member),
                                ];
                            }
                        }

                        break 2;
                    }
                }
            }
        }

        $this->rc->output->command('plugin.expand_attendee_callback', $result);
    }

    /**
     * Merge attendees of the old and new event version
     * with keeping current user and his delegatees status
     *
     * @param array &$new   New object data
     * @param array $old    Old object data
     * @param bool  $status New status of the current user
     */
    public function merge_attendees(&$new, $old, $status = null)
    {
        if (empty($status)) {
            $emails    = $this->get_user_emails();
            $delegates = [];
            $attendees = [];

            // keep attendee status of the current user
            foreach ((array) $new['attendees'] as $i => $attendee) {
                if (empty($attendee['email'])) {
                    continue;
                }

                $attendees[] = $email = strtolower($attendee['email']);

                if (in_array($email, $emails)) {
                    foreach ($old['attendees'] as $_attendee) {
                        if ($attendee['email'] == $_attendee['email']) {
                            $new['attendees'][$i] = $_attendee;
                            if ($_attendee['status'] == 'DELEGATED' && ($email = $_attendee['delegated-to'])) {
                                $delegates[] = strtolower($email);
                            }

                            break;
                        }
                    }
                }
            }

            // make sure delegated attendee is not lost
            foreach ($delegates as $delegatee) {
                if (!in_array($delegatee, $attendees)) {
                    foreach ((array) $old['attendees'] as $attendee) {
                        if ($attendee['email'] && ($email = strtolower($attendee['email'])) && $email == $delegatee) {
                            $new['attendees'][] = $attendee;
                            break;
                        }
                    }
                }
            }
        }

        // We also make sure that status of any attendee
        // is not overriden by NEEDS-ACTION if it was already set
        // which could happen if you work with shared events
        foreach ((array) $new['attendees'] as $i => $attendee) {
            if ($attendee['email'] && ($attendee['status'] ?? '') == 'NEEDS-ACTION') {
                foreach ($old['attendees'] as $_attendee) {
                    if ($attendee['email'] == $_attendee['email']) {
                        $new['attendees'][$i]['status'] = $_attendee['status'];
                        unset($new['attendees'][$i]['rsvp']);
                        break;
                    }
                }
            }
        }
    }


    /*********  Static utility functions  *********/

    /**
     * Convert the internal structured data into a vcalendar rrule 2.0 string
     */
    public static function to_rrule($recurrence, $allday = false)
    {
        if (is_string($recurrence)) {
            return $recurrence;
        }

        $rrule = '';
        foreach ((array)$recurrence as $k => $val) {
            $k = strtoupper($k);
            switch ($k) {
                case 'UNTIL':
                    // convert to UTC according to RFC 5545
                    if (is_a($val, 'DateTime')) {
                        if (!$allday && empty($val->_dateonly)) {
                            $until = clone $val;
                            $until->setTimezone(new DateTimeZone('UTC'));
                            $val = $until->format('Ymd\THis\Z');
                        } else {
                            $val = $val->format('Ymd');
                        }
                    }
                    break;
                case 'RDATE':
                case 'EXDATE':
                    foreach ((array)$val as $i => $ex) {
                        if (is_a($ex, 'DateTime')) {
                            $val[$i] = $ex->format('Ymd\THis');
                        }
                    }
                    $val = implode(',', (array)$val);
                    break;
                case 'EXCEPTIONS':
                    continue 2;
            }

            if (strlen($val)) {
                $rrule .= $k . '=' . $val . ';';
            }
        }

        return rtrim($rrule, ';');
    }

    /**
     * Convert from fullcalendar date format to PHP date() format string
     */
    public static function to_php_date_format($from)
    {
        if (!is_string($from)) {
            return '';
        }

        // "dd.MM.yyyy HH:mm:ss" => "d.m.Y H:i:s"
        return strtr(strtr($from, [
            'YYYY' => 'Y',
            'YY'   => 'y',
            'yyyy' => 'Y',
            'yy'   => 'y',
            'MMMM' => 'F',
            'MMM'  => 'M',
            'MM'   => 'm',
            'M'    => 'n',
            'dddd' => 'l',
            'ddd'  => 'D',
            'DD'   => 'd',
            'D'    => 'j',
            'HH'   => '**',
            'hh'   => '%%',
            'H'    => 'G',
            'h'    => 'g',
            'mm'   => 'i',
            'ss'   => 's',
            'TT'   => 'A',
            'tt'   => 'a',
            'T'    => 'A',
            't'    => 'a',
            'u'    => 'c',
        ]), [
            '**'   => 'H',
            '%%'   => 'h',
        ]);
    }

    /**
     * Convert from PHP date() format to fullcalendar (MomentJS) format string
     */
    public static function from_php_date_format($from)
    {
        if (!is_string($from)) {
            return '';
        }

        // "d.m.Y H:i:s" => "dd.MM.yyyy HH:mm:ss"
        return strtr($from, [
            'y' => 'YY',
            'Y' => 'YYYY',
            'M' => 'MMM',
            'F' => 'MMMM',
            'm' => 'MM',
            'n' => 'M',
            'j' => 'D',
            'd' => 'DD',
            'D' => 'ddd',
            'l' => 'dddd',
            'H' => 'HH',
            'h' => 'hh',
            'G' => 'H',
            'g' => 'h',
            'i' => 'mm',
            's' => 'ss',
            'c' => '',
        ]);
    }
}
