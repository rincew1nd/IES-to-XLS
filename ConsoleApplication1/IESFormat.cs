﻿// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
// 
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
// 
// You should have received a copy of the GNU General Public License
// along with this program.  If not, see <http://www.gnu.org/licenses/>.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace IESConverter
	{
		public class IesFile : IDisposable
		{
			private Stream _stream;
			private BinaryReader _reader;
			private byte _xorKey;

			public List<IesColumn> Columns { get; private set; }
			public IesHeader Header { get; private set; }
			public List<IesRow> Rows { get; private set; }

			public IesFile(Stream stream)
			{
				_stream = stream;
				_reader = new BinaryReader(_stream);
				_xorKey = 1;

				this.ReadHeader();
				this.ReadColumns();
				this.ReadRows();
			}

			public IesFile(byte[] content)
				: this(new MemoryStream(content))
			{
			}

			private string DecryptString(byte[] data, Encoding encoding = null)
			{
				if (encoding == null)
					encoding = Encoding.UTF8;

				var bytes = new byte[data.Length];
				for (var i = 0; i < data.Length; i++)
					bytes[i] = (byte)(data[i] ^ _xorKey);

				return encoding.GetString(bytes).TrimEnd(new char[] { '\x0001' });
			}

			public void Dispose()
			{
				_reader?.Close();
			}

			private void ReadColumns()
			{
				_stream.Seek((-((long)this.Header.ResourceOffset) - this.Header.DataOffset), SeekOrigin.End);

				this.Columns = new List<IesColumn>();
				for (var i = 0; i < this.Header.ColumnCount; i++)
				{
					var item = new IesColumn
					{
						Name = this.DecryptString(_reader.ReadBytes(0x40), null),
						Name2 = this.DecryptString(_reader.ReadBytes(0x40), null),
						Type = (ColumnType) _reader.ReadUInt16()
					};
					_reader.ReadUInt32();
					item.Position = _reader.ReadUInt16();

					// Duplicates
					var old = item.Name;
					for (var j = 1; this.Columns.Exists(a => a.Name == item.Name); ++j)
						item.Name = old + "_" + j;

					this.Columns.Add(item);
				}
				this.Columns.Sort();
			}

			private void ReadHeader()
			{
				this.Header = new IesHeader
				{
					Name = Encoding.UTF8.GetString(_reader.ReadBytes(0x80))
				};
				_reader.ReadUInt32();
				this.Header.DataOffset = _reader.ReadUInt32();
				this.Header.ResourceOffset = _reader.ReadUInt32();
				this.Header.FileSize = _reader.ReadUInt32();
				_reader.ReadUInt16();
				this.Header.RowCount = _reader.ReadUInt16();
				this.Header.ColumnCount = _reader.ReadUInt16();
				this.Header.NumberColumnCount = _reader.ReadUInt16();
				this.Header.StringColumnCount = _reader.ReadUInt16();
				_reader.ReadUInt16();
			}

			private void ReadRows()
			{
				_reader.BaseStream.Seek(-((long)this.Header.ResourceOffset), SeekOrigin.End);

				this.Rows = new List<IesRow>();
				for (var i = 0; i < this.Header.RowCount; ++i)
				{
					_reader.ReadUInt32();
					var count = _reader.ReadUInt16();
					_reader.ReadBytes(count);

					var item = new IesRow();
					for (var j = 0; j < this.Columns.Count; ++j)
					{
						var column = this.Columns[j];

						if (column.IsNumber)
						{
							var nan = _reader.ReadSingle();
							var maxValue = uint.MaxValue;
							try
							{
								maxValue = (uint)nan;
							}
							catch
							{
								nan = float.NaN;
							}

							item.Add(
								column.Name,
								Math.Abs((float) (nan - maxValue)) < float.Epsilon ?
									maxValue : nan
							);
						}
						else
						{
							var length = _reader.ReadUInt16();
							var str = "";
							if (length > 0)
								str = this.DecryptString(_reader.ReadBytes(length), null);

							item.Add(column.Name, str);
						}
					}

					this.Rows.Add(item);
					_reader.BaseStream.Seek((long)this.Header.StringColumnCount, SeekOrigin.Current);
				}
			}
		}

		public class IesHeader
		{
			public uint DataOffset { get; set; }
			public uint ResourceOffset { get; set; }
			public uint FileSize { get; set; }
			public string Name { get; set; }
			public ushort ColumnCount { get; set; }
			public ushort RowCount { get; set; }
			public ushort NumberColumnCount { get; set; }
			public ushort StringColumnCount { get; set; }
		}

		public class IesColumn : IComparable<IesColumn>
		{
			public string Name { get; set; }
			public string Name2 { get; set; }
			public ColumnType Type { get; set; }
			public ushort Position { get; set; }

			public bool IsNumber => (this.Type == ColumnType.Float);

			public int CompareTo(IesColumn other)
			{
				if (((this.Type == other.Type) ||
					((this.Type == ColumnType.String) &&
					(other.Type == ColumnType.String2))) ||
					((this.Type == ColumnType.String2) &&
					(other.Type == ColumnType.String)))
					return this.Position.CompareTo(other.Position);

				if (this.Type < other.Type)
					return -1;

				return 1;
			}
		}

		public enum ColumnType
		{
			Float,
			String,
			String2
		}

		public class IesRow : Dictionary<string, object>
		{
			public float GetFloat(string name)
			{
				if (!ContainsKey(name))
					throw new ArgumentException("Unknown field: " + name);

				if (this[name] is float) return (float)this[name];
				if (this[name] is uint) return (float)(uint)this[name];

				throw new ArgumentException(name + " is not numeric");
			}

			public uint GetUInt(string name)
			{
				return (uint)GetInt(name);
			}

			public int GetInt(string name)
			{
				if (!ContainsKey(name))
					throw new ArgumentException("Unknown field: " + name);

				if (this[name] is float) return (int)(float)this[name];
				if (this[name] is uint) return (int)(uint)this[name];

				throw new ArgumentException(name + " is not numeric");
			}

			public string GetString(string name)
			{
				if (!ContainsKey(name))
					throw new ArgumentException("Unknown field: " + name);

				if (this[name] is string)
					return (string)this[name];

				throw new ArgumentException(name + " is not a string");
			}
		}
	}
