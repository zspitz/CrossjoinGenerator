﻿using System.Diagnostics.CodeAnalysis;
using System.Runtime.CompilerServices;
using System.Windows;

namespace Util;

public static class Functions {
    private static string parseCallerMemberName(string? callerMemberName) {
        ArgumentNullException.ThrowIfNull(callerMemberName);
        var s = callerMemberName;
        if (s.EndsWith("Property")) { s = s[..^"Property".Length]; }
        return s;
    }

    public static DependencyProperty DPRegister<TProperty, TOwner>(
        [AllowNull] TProperty defaultValue = default,
        [CallerMemberName] string? propertyName = null
    ) =>
        DependencyProperty.Register(parseCallerMemberName(propertyName), typeof(TProperty), typeof(TOwner), new PropertyMetadata(defaultValue, null));

    public static DependencyProperty DPRegister<TProperty, TOwner>(
        TProperty defaultValue,
        Action<TOwner, TProperty, TProperty> callback,
        [CallerMemberName] string? propertyName = null
    ) where TOwner : DependencyObject =>
        DependencyProperty.Register(
            parseCallerMemberName(propertyName),
            typeof(TProperty),
            typeof(TOwner),
            new PropertyMetadata(defaultValue, (d, e) =>
                callback((TOwner)d, (TProperty)e.NewValue, (TProperty)e.OldValue)
            )
        );

    public static DependencyProperty DPRegister<TProperty, TOwner>(
        TProperty defaultValue,
        FrameworkPropertyMetadataOptions flags,
        [CallerMemberName] string? propertyName = null
    ) =>
        DependencyProperty.Register(parseCallerMemberName(propertyName), typeof(TProperty), typeof(TOwner), new FrameworkPropertyMetadata(defaultValue, flags, null));

    [Obsolete("Use DPRegister overload which explicitly casts the owner, new and old values.")]
    public static DependencyProperty DPRegister<TProperty, TOwner>(
        TProperty defaultValue,
        FrameworkPropertyMetadataOptions flags,
        PropertyChangedCallback? callback,
        [CallerMemberName] string? propertyName = null
    ) =>
        DependencyProperty.Register(parseCallerMemberName(propertyName), typeof(TProperty), typeof(TOwner), new FrameworkPropertyMetadata(defaultValue, flags, callback));

    public static DependencyProperty DPRegister<TProperty, TOwner>(
        TProperty defaultValue,
        FrameworkPropertyMetadataOptions flags,
        Action<TOwner, TProperty, TProperty> callback,
        [CallerMemberName] string? propertyName = null
    ) where TOwner : DependencyObject =>
        DependencyProperty.Register(
            parseCallerMemberName(propertyName),
            typeof(TProperty),
            typeof(TOwner),
            new FrameworkPropertyMetadata(defaultValue, flags, (d, e) =>
                callback((TOwner)d, (TProperty)e.NewValue, (TProperty)e.OldValue)
            )
        );

    public static DependencyProperty DPRegisterAttached<TProperty, TOwner>([AllowNull] TProperty defaultValue = default, PropertyChangedCallback? callback = null, [CallerMemberName] string? propertyName = null) =>
        DependencyProperty.RegisterAttached(parseCallerMemberName(propertyName), typeof(TProperty), typeof(TOwner), new PropertyMetadata(defaultValue, callback));

    public static DependencyProperty DPRegisterAttached<TProperty, TOwner>(
        TProperty defaultValue,
        FrameworkPropertyMetadataOptions flags,
        Action<TOwner, TProperty, TProperty> callback,
        [CallerMemberName] string? propertyName = null
    ) where TOwner : DependencyObject =>
        DependencyProperty.RegisterAttached(
            parseCallerMemberName(propertyName),
            typeof(TProperty),
            typeof(TOwner),
            new FrameworkPropertyMetadata(defaultValue, flags, (d, e) =>
                callback((TOwner)d, (TProperty)e.NewValue, (TProperty)e.OldValue)
            )
        );

    public static DependencyProperty DPRegisterAttached<TProperty, TOwner>(FrameworkPropertyMetadataOptions flags, PropertyChangedCallback? callback = null, [CallerMemberName] string? propertyName = null) =>
        DependencyProperty.RegisterAttached(parseCallerMemberName(propertyName), typeof(TProperty), typeof(TOwner), new FrameworkPropertyMetadata(DependencyProperty.UnsetValue, flags, callback));

    public static bool IsEqual<T>(T current, T newValue) => EqualityComparer<T>.Default.Equals(current, newValue);

    public static T IIFE<T>(Func<T> fn) => fn();
}

