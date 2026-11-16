// Copyright (c) 2026 Erik Witte. See License.txt in the repository root.

using System.Text.Json.Serialization;

namespace STFM.Extension.Screening;

// No reflection dependency: supported by MAUI trimming and AOT.
[JsonSourceGenerationOptions(PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase,
    WriteIndented = true, UseStringEnumConverter = true)]
[JsonSerializable(typeof(ScreeningExamination))]
internal partial class ScreeningJsonContext : JsonSerializerContext;
