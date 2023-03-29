using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;


namespace BBS.ST.BHC.PDC.AutoUpdater.Security
{
  /// <summary>
  ///   The AesEncoder class provides methods for en and decrypting data using
  ///   256 bit AES (Advanced Encryption Standard) (en)decryption with a block size of 256.
  /// </summary>
  /// <remarks>
  ///   For more Informations about en/decryption take a look at:
  ///   http://www.developerfusion.co.uk/show/4647/4/
  /// </remarks>
  public sealed class AesEncoder
  {
    private static SymmetricAlgorithm   MySymmetricAlgorithm = null;
    private static SHA256               MySHA = null;


    // .ctor and .cctor
    // ////////////////////////////////////////////////////////////////////////////////////////////

    #region .ctor and .cctor

    /// <summary>
    ///   Creates a new instance of the Cipher class.
    /// </summary>
    private AesEncoder()
    {
      this.InitializeMyClass();
    }

    /// <summary>
    ///   Initializes the static SymmetricAlgorithm.
    /// </summary>
    static AesEncoder()
    {
      AesEncoder.MySymmetricAlgorithm = new RijndaelManaged();

      AesEncoder.MySHA = SHA256.Create();


      AesEncoder.MySymmetricAlgorithm.BlockSize = 256;

      AesEncoder.MySymmetricAlgorithm.KeySize = 256;

      AesEncoder.MySymmetricAlgorithm.Padding = PaddingMode.PKCS7;

      AesEncoder.MySymmetricAlgorithm.IV = MySHA.ComputeHash(Encoding.UTF8.GetBytes("a3eCT675yX"));
    }

    #endregion


    // Clean up
    // ////////////////////////////////////////////////////////////////////////////////////////////


    // Initialize
    // ////////////////////////////////////////////////////////////////////////////////////////////

    #region InitializeMyClass

    /// <summary>
    ///   Initializes the new instance.
    /// </summary>
    private void InitializeMyClass()
    {
    }

    #endregion


    // Methods
    // ////////////////////////////////////////////////////////////////////////////////////////////

    #region Decrypt

    /// <summary>
    ///   Decrypts the specifies encrypted text (Base64 format) using the specified key with a
    ///   256 bit AES (Advanced Encryption Standard) Decryption with a block size of 256.
    /// </summary>
    /// <param name="text">
    ///   The encrypted text (Base64 format) to decrypt.
    /// </param>
    /// <param name="key">
    ///   The key to use for decryption.
    /// </param>
    /// <returns>
    ///   The decrypted plan text.
    /// </returns>
    /// <exception cref="CryptographicException">
    ///   Decrypting encrypted text with a key that differs from the key used for encrypting
    ///   the text may be cause an CryptographicException.
    /// </exception>
    public static String Decrypt(String text, String key)
    {
      Byte[]              input;
      Byte[]              output;


      if (String.IsNullOrEmpty(text)) return String.Empty;

      if (String.IsNullOrEmpty(key)) return text;


      AesEncoder.MySymmetricAlgorithm.Key = MySHA.ComputeHash(Encoding.UTF8.GetBytes(key));


      input = Convert.FromBase64String(text);

      output = Decrypt(input);


      if(output == null || output.Length == 0) return String.Empty;


      return Encoding.UTF8.GetString(output);
    }

    /// <summary>
    ///   Decryptes the specified byte array.
    /// </summary>
    /// <param name="input">
    ///   The byte array to decrypt.
    /// </param>
    /// <returns>
    ///   The decrypted byte array.
    /// </returns>
    private static Byte[] Decrypt(Byte[] input)
    {
      Byte[]            buffer = null;
      CryptoStream      cryptoStream = null;
      Int32             length;
      MemoryStream      memoryStream = null;
      Byte[]            output = null;
      ICryptoTransform  transform = null;


      buffer = new Byte[input.Length];

      try
      {
        memoryStream = new MemoryStream(input);

        transform = AesEncoder.MySymmetricAlgorithm.CreateDecryptor();

        cryptoStream = new CryptoStream(memoryStream, transform, CryptoStreamMode.Read);
        
        length = cryptoStream.Read(buffer, 0, input.Length);


        output = new Byte[length];

        Array.Copy(buffer, output, length);
      }
      finally
      {
        if (cryptoStream != null) cryptoStream.Close();

        if (transform != null) transform.Dispose();
      
        if (memoryStream != null) memoryStream.Close();
      }      


      return output;
    }

    #endregion


    #region Encrypt

    /// <summary>
    ///   Encrypts the specified plan text using the specified key with a
    ///   256 bit AES (Advanced Encryption Standard) Encryption with a block size of 256.
    /// </summary>
    /// <param name="text">
    ///   The plan text to encrypt.
    /// </param>
    /// <param name="key">
    ///   The key to use for encryption. 
    ///   The key will be encoded by UFT8.
    /// </param>
    /// <returns>
    ///   The encrypted text (Base64 Format).
    /// </returns>
    public static String Encrypt(String text, String key)
    {
      Byte[]              input;
      Byte[]              output;


      if (String.IsNullOrEmpty(text)) return String.Empty;

      if (String.IsNullOrEmpty(key)) return text;

      
      AesEncoder.MySymmetricAlgorithm.Key = MySHA.ComputeHash(Encoding.UTF8.GetBytes(key));
      

      input = Encoding.UTF8.GetBytes(text);

      output = Encrypt(input);


      if(output == null || output.Length == 0) return String.Empty;


      return Convert.ToBase64String(output);
    }

    /// <summary>
    ///   Encryptes the specified byte array.
    /// </summary>
    /// <param name="input">
    ///   The byte array to encrypt.
    /// </param>
    /// <returns>
    ///   The encrypted byte array.
    /// </returns>
    private static Byte[] Encrypt(Byte[] input)
    {
      CryptoStream      cryptoStream = null;
      MemoryStream      memoryStream = null;
      Byte[]            output = null;
      ICryptoTransform  transform = null;


      try
      {
        memoryStream = new MemoryStream();

        transform = AesEncoder.MySymmetricAlgorithm.CreateEncryptor();

        cryptoStream = new CryptoStream(memoryStream, transform, CryptoStreamMode.Write);
        

        cryptoStream.Write(input, 0, input.Length);

        cryptoStream.FlushFinalBlock();

        output = memoryStream.ToArray();
      }
      finally
      {
        if (cryptoStream != null) cryptoStream.Close();

        if (transform != null) transform.Dispose();
        
        if (memoryStream != null) memoryStream.Close();
      }


      return output;
    }


    #endregion


    // Properties
    // ////////////////////////////////////////////////////////////////////////////////////////////
  }
}