import datetime
import json
from decimal import Decimal
from typing import Any, Callable, Iterable, Union

from django.conf import settings as django_settings
from django.utils.dateparse import parse_date, parse_datetime, parse_time
from openpyxl.cell import Cell
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles import NamedStyle
from openpyxl.styles.numbers import (
	FORMAT_DATE_DATETIME,
	FORMAT_DATE_TIME4,
	FORMAT_DATE_YYYYMMDD2,
	FORMAT_NUMBER,
	FORMAT_NUMBER_00,
)
from openpyxl.worksheet.worksheet import Worksheet
from rest_framework import ISO_8601
from rest_framework.fields import (
	DateField,
	DateTimeField,
	DecimalField,
	Field,
	FloatField,
	IntegerField,
	TimeField,
)
from rest_framework.settings import api_settings as drf_settings

ESCAPE_CHARS = ('=', '-', '+', '@', '\t', '\r', '\n',)


def get_setting(key):
	return getattr(django_settings, 'DRF_RENDERER_XLSX_' + key, None)


class XLSXField:
	sanitize = True

	def __init__(self, key, value, field: Field, style: NamedStyle, mapping: Union[str, Callable]):
		self.key = key
		self.original_value = value
		self.drf_field = field
		self.style = style or NamedStyle()
		self.mapping = mapping
		self.value = self.init_value(value)

	def init_value(self, value):
		return value

	def sanitize_value(self, value):
		# prepend ' if value is starting with possible malicious char
		if self.sanitize and value:
			str_value = str(value)
			str_value = ILLEGAL_CHARACTERS_RE.sub('', str_value)  # remove ILLEGAL_CHARACTERS so it doesn't crash
			return "'" + str_value if str_value.startswith(ESCAPE_CHARS) else str_value
		return value

	def custom_mapping(self):
		if type(self.mapping) is str:
			return self.value.get(self.mapping)
		elif callable(self.mapping):
			return self.mapping(self.value)
		return self.value

	def prep_value(self) -> Any:
		return self.value

	def prep_cell(self, cell: Cell):
		# We cannot apply the whole style directly, otherwise we cannot override any part of it
		cell.font = self.style.font
		cell.fill = self.style.fill
		cell.alignment = self.style.alignment
		cell.border = self.style.border
		cell.number_format = self.style.number_format

	def cell(self, ws: Worksheet, row, column) -> Cell:
		# If we have a custom mapping use it and done. If not prep value for output
		value = self.sanitize_value(self.custom_mapping() if self.mapping else self.prep_value())
		cell: Cell = ws.cell(row, column, value)
		self.prep_cell(cell)
		return cell


class XLSXNumberField(XLSXField):
	sanitize = False

	def __init__(self, number_format, **kwargs):
		self.number_format = number_format
		super().__init__(**kwargs)

	def init_value(self, value):
		try:
			if isinstance(self.drf_field, IntegerField) and type(value) != int:
				return int(value)
			elif isinstance(self.drf_field, FloatField) and type(value) != float:
				return float(value)
			elif isinstance(self.drf_field, DecimalField) and type(value) != Decimal:
				return Decimal(value)
		except:
			pass
		return value

	def prep_cell(self, cell: Cell):
		super().prep_cell(cell)
		if self.number_format:
			cell.number_format = self.number_format
		elif isinstance(self.drf_field, IntegerField):
			cell.number_format = get_setting('INTEGER_FORMAT') or FORMAT_NUMBER
		else:
			cell.number_format = get_setting('DECIMAL_FORMAT') or FORMAT_NUMBER_00


class XLSXDateField(XLSXField):
	sanitize = False

	def __init__(self, date_format, **kwargs):
		self.date_format = date_format
		super().__init__(**kwargs)

	def _parse_date(self, value, setting_format, iso_parse_func):
		# Parse format is DRF output format: DATETIME_FORMAT, DATE_FORMAT or TIME_FORMAT
		parse_format = getattr(drf_settings, setting_format)
		# Use the provided iso parse function for the special case ISO_8601
		if parse_format.lower() == ISO_8601:
			return iso_parse_func(value)
		else:
			return datetime.datetime.strptime(value, parse_format)

	def init_value(self, value):
		try:
			if isinstance(self.drf_field, DateTimeField) and type(value) != datetime.datetime:
				return self._parse_date(value, 'DATETIME_FORMAT', parse_datetime)
			elif isinstance(self.drf_field, FloatField) and type(value) != datetime.date:
				return self._parse_date(value, 'DATE_FORMAT', parse_date)
			elif isinstance(self.drf_field, DecimalField) and type(value) != datetime.time:
				return self._parse_date(value, 'TIME_FORMAT', parse_time)
		except:
			return value

	def prep_cell(self, cell: Cell):
		super().prep_cell(cell)
		if self.date_format:
			cell.number_format = self.date_format
		elif isinstance(self.drf_field, DateTimeField):
			cell.number_format = get_setting('DATETIME_FORMAT') or FORMAT_DATE_DATETIME
		elif isinstance(self.drf_field, DateField):
			cell.number_format = get_setting('DATE_FORMAT') or FORMAT_DATE_YYYYMMDD2
		elif isinstance(self.drf_field, TimeField):
			cell.number_format = get_setting('TIME_FORMAT') or FORMAT_DATE_TIME4


class XLSXListField(XLSXField):
	def __init__(self, list_sep, **kwargs):
		self.list_sep = list_sep or ", "
		super().__init__(**kwargs)

	def prep_value(self) -> Any:
		if len(self.value) > 0 and isinstance(self.value[0], Iterable):
			# array of array; write as json
			return json.dumps(self.value)
		else:
			# Flatten the array into a comma separated string to fit
			# in a single spreadsheet column
			return self.list_sep.join(map(str, self.value))


class XLSXBooleanField(XLSXField):
	sanitize = False

	def __init__(self, boolean_display, **kwargs):
		self.boolean_display = boolean_display
		super().__init__(**kwargs)

	def prep_value(self) -> Any:
		if self.boolean_display:
			return str(self.boolean_display.get(self.value, self.value))
		return self.value
