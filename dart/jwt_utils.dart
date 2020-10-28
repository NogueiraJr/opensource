import 'package:jaguar_jwt/jaguar_jwt.dart';
import 'package:projeto/model/user_model.dart';

class JwtUtils {
  static const String _jwtChavePrivada =
      'VrjXkqxXpvut607^Xpvut607^tamA%*Xpvut607^M&SeXpvut607^sPD6va';
  static String gerarTokenJwt(UsuarioModel usuario) {
    final claimSet = JwtClaim(
        issuer: 'http://localhost',
        subject: usuario.id.toString(),
        otherClaims: <String, dynamic>{},
        maxAge: const Duration(days: 2));

    final token = 'Bearer ${issueJwtHS256(claimSet, _jwtChavePrivada)}';

    return token;
  }

  static JwtClaim verificarToken(String token) {
    return verifyJwtHS256Signature(token, _jwtChavePrivada);
  }
}
