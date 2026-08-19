//! Zero-dependency newtypes for the ABI's three temporal column types.
//!
//! `XL_T_DATE`/`XL_T_TIME`/`XL_T_TIMESTAMP` each have an exact wire representation (see
//! `excelreader.h`), and these types are that representation with a name attached - no conversion,
//! no calendar arithmetic, nothing to get wrong. The C++ wrapper maps the same three columns onto
//! `std::chrono`; Rust's standard library has no calendar type, so the crate carries these instead
//! and offers `chrono` interop behind the optional `chrono` feature.

/// Days since 1970-01-01. The wire form of an `XL_T_DATE` column.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct Date {
    pub days_since_epoch: i32,
}

/// Microseconds since midnight. The wire form of an `XL_T_TIME` column.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct Time {
    pub micros_since_midnight: i64,
}

/// Microseconds since 1970-01-01T00:00:00Z. The wire form of an `XL_T_TIMESTAMP` column.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct Timestamp {
    pub micros_since_epoch: i64,
}

impl Date {
    #[must_use]
    pub fn new(days_since_epoch: i32) -> Self {
        Self { days_since_epoch }
    }
}

impl Time {
    #[must_use]
    pub fn new(micros_since_midnight: i64) -> Self {
        Self {
            micros_since_midnight,
        }
    }
}

impl Timestamp {
    #[must_use]
    pub fn new(micros_since_epoch: i64) -> Self {
        Self { micros_since_epoch }
    }
}

/// `chrono` interop. Each conversion is total over every value a real workbook can produce, and
/// panics only on a value outside `chrono`'s own representable range - which the native converters
/// cannot emit, since Excel's serial dates are bounded far inside it.
#[cfg(feature = "chrono")]
mod chrono_interop {
    use super::{Date, Time, Timestamp};
    use chrono::{DateTime, Datelike, NaiveDate, NaiveDateTime, NaiveTime, Timelike};

    /// `NaiveDate::from_num_days_from_ce_opt` counts from 0001-01-01; 1970-01-01 sits 719_163 days
    /// after it. Going through this offset (rather than adding a `Days` duration to a constructed
    /// epoch date) keeps the conversion a single addition with no intermediate `Option` to unwrap.
    const DAYS_CE_TO_UNIX_EPOCH: i32 = 719_163;

    const MICROS_PER_SEC: i64 = 1_000_000;
    const NANOS_PER_MICRO: u32 = 1_000;

    impl From<Date> for NaiveDate {
        fn from(value: Date) -> Self {
            NaiveDate::from_num_days_from_ce_opt(value.days_since_epoch + DAYS_CE_TO_UNIX_EPOCH)
                .expect("XL_T_DATE value is outside chrono::NaiveDate's representable range")
        }
    }

    impl From<NaiveDate> for Date {
        fn from(value: NaiveDate) -> Self {
            Date::new(value.num_days_from_ce() - DAYS_CE_TO_UNIX_EPOCH)
        }
    }

    impl From<Time> for NaiveTime {
        fn from(value: Time) -> Self {
            let secs = value.micros_since_midnight.div_euclid(MICROS_PER_SEC);
            let micros = value.micros_since_midnight.rem_euclid(MICROS_PER_SEC);
            NaiveTime::from_num_seconds_from_midnight_opt(
                u32::try_from(secs)
                    .expect("XL_T_TIME value is outside chrono::NaiveTime's representable range"),
                micros as u32 * NANOS_PER_MICRO,
            )
            .expect("XL_T_TIME value is outside chrono::NaiveTime's representable range")
        }
    }

    impl From<NaiveTime> for Time {
        fn from(value: NaiveTime) -> Self {
            let secs = i64::from(value.num_seconds_from_midnight());
            let micros = i64::from(value.nanosecond() / NANOS_PER_MICRO);
            Time::new(secs * MICROS_PER_SEC + micros)
        }
    }

    impl From<Timestamp> for NaiveDateTime {
        fn from(value: Timestamp) -> Self {
            DateTime::from_timestamp_micros(value.micros_since_epoch)
                .expect(
                    "XL_T_TIMESTAMP value is outside chrono::NaiveDateTime's representable range",
                )
                .naive_utc()
        }
    }

    impl From<NaiveDateTime> for Timestamp {
        fn from(value: NaiveDateTime) -> Self {
            Timestamp::new(value.and_utc().timestamp_micros())
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn newtypes_carry_the_wire_value_verbatim() {
        assert_eq!(Date::new(-1).days_since_epoch, -1);
        assert_eq!(Time::new(0).micros_since_midnight, 0);
        assert_eq!(Timestamp::new(i64::MIN).micros_since_epoch, i64::MIN);
    }

    #[cfg(feature = "chrono")]
    mod chrono {
        use super::*;
        use ::chrono::{NaiveDate, NaiveDateTime, NaiveTime};

        #[test]
        fn date_zero_is_the_unix_epoch() {
            let date: NaiveDate = Date::new(0).into();
            assert_eq!(date, NaiveDate::from_ymd_opt(1970, 1, 1).unwrap());
        }

        #[test]
        fn date_round_trips_through_chrono_in_both_directions() {
            for days in [-25_567, -1, 0, 1, 19_000, 2_958_465 - 25_569] {
                let date: NaiveDate = Date::new(days).into();
                assert_eq!(Date::from(date), Date::new(days), "days = {days}");
            }
        }

        #[test]
        fn time_keeps_microsecond_precision() {
            let time: NaiveTime = Time::new(13 * 3_600_000_000 + 45 * 60_000_000 + 1).into();
            assert_eq!(time, NaiveTime::from_hms_micro_opt(13, 45, 0, 1).unwrap());
            assert_eq!(
                Time::from(time),
                Time::new(13 * 3_600_000_000 + 45 * 60_000_000 + 1)
            );
        }

        #[test]
        fn timestamp_round_trips_across_the_epoch() {
            for micros in [-1_000_000, -1, 0, 1, 1_700_000_000_000_000] {
                let stamp: NaiveDateTime = Timestamp::new(micros).into();
                assert_eq!(
                    Timestamp::from(stamp),
                    Timestamp::new(micros),
                    "micros = {micros}"
                );
            }
        }
    }
}
