using System;
using Microsoft.VisualStudio.Shell;

[AttributeUsage(AttributeTargets.Assembly, AllowMultiple = true)]
public sealed class ProvideGeneratorAttribute : RegistrationAttribute
{
    private readonly Type _generatorType;
    private readonly Guid _languageServiceGuid;
    private string _name;
    private string _description;
    private bool _generatesDesignTimeSource;

    public ProvideGeneratorAttribute(Type generatorType, string languageServiceGuid)
    {
        if (generatorType == null)
            throw new ArgumentNullException("generatorType");
        if (languageServiceGuid == null)
            throw new ArgumentNullException("languageServiceGuid");
        if (string.IsNullOrEmpty(languageServiceGuid))
            throw new ArgumentException("languageServiceGuid cannot be empty");

        _generatorType = generatorType;
        _languageServiceGuid = new Guid(languageServiceGuid);
        _name = _generatorType.Name;
    }

    public Type GeneratorType
    {
        get
        {
            return _generatorType;
        }
    }

    public Guid LanguageServiceGuid
    {
        get
        {
            return _languageServiceGuid;
        }
    }

    public string Name
    {
        get
        {
            return _name;
        }

        set
        {
            if (value == null)
                throw new ArgumentNullException("value");
            if (string.IsNullOrEmpty(value))
                throw new ArgumentException("value cannot be empty");

            _name = value;
        }
    }

    public string Description
    {
        get
        {
            return _description;
        }

        set
        {
            _description = value;
        }
    }

    public bool GeneratesDesignTimeSource
    {
        get
        {
            return _generatesDesignTimeSource;
        }

        set
        {
            _generatesDesignTimeSource = value;
        }
    }

    private string RegistrationKey
    {
        get
        {
            return string.Format(@"Generators\{0}\{1}", LanguageServiceGuid.ToString("B"), Name);
        }
    }

    public override void Register(RegistrationContext context)
    {
        using (Key key = context.CreateKey(RegistrationKey))
        {
            if (!string.IsNullOrEmpty(Description))
                key.SetValue(string.Empty, Description);
            key.SetValue("CLSID", GeneratorType.GUID.ToString("B"));
            key.SetValue("GeneratesDesignTimeSource", GeneratesDesignTimeSource ? 1 : 0);
        }
    }

    public override void Unregister(RegistrationContext context)
    {
        context.RemoveKey(RegistrationKey);
    }
}